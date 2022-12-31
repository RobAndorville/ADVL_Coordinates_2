Imports System.Reflection.Emit
Imports System.Runtime
Imports System.Security.Cryptography
Imports ADVL_Utilities_Library_1
Imports Microsoft.Runtime

Public Class frmConversion
    'Convert a coordinate from one reference system to another.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    Dim WithEvents Conversion As New clsCoordTransformation

    Dim CrsInputSearch As DataSet = New DataSet
    Dim CrsOutputSearch As DataSet = New DataSet

    'NOTE: The splitter distances are not restored correctly if the tab does not have focus. 'The distances are saved separately here.
    'Dim InputSplitDist1 As Integer
    'Dim InputSplitDist2 As Integer
    'Dim OutputSplitDist1 As Integer
    'Dim OutputSplitDist2 As Integer
    Dim SplitDist1, SplitDist2, SplitDist3, SplitDist4, SplitDist5, SplitDist6, SplitDist7, SplitDist8, SplitDist9 As Integer
    Dim SplitDist10, SplitDist11, SplitDist12, SplitDist13, SplitDist14, SplitDist15 As Integer
    Dim DirectCoordOpCode As Integer = -1
    Dim InputToWgs84CoordOpCode As Integer = -1
    Dim Wgs84ToOutputCoordOpCode As Integer = -1


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
            Conversion.EpsgDatabasePath = _epsgDatabasePath
        End Set
    End Property

    Private _selInputRowNo As Integer = -1 'The selected row number in the list of Input coordinate reference systems.
    Property SelInputRowNo As Integer
        Get
            Return _selInputRowNo
        End Get
        Set(value As Integer)
            _selInputRowNo = value
        End Set
    End Property

    Private _selOutputRowNo As Integer = -1 'The selected row number in the list of Output coordinate reference systems.
    Property SelOutputRowNo As Integer
        Get
            Return _selOutputRowNo
        End Get
        Set(value As Integer)
            _selOutputRowNo = value
        End Set
    End Property

    Private _modified As Boolean = False 'If True, the conversion settings have been modified.
    Property Modified As Boolean
        Get
            Return _modified
        End Get
        Set(value As Boolean)
            _modified = value
        End Set
    End Property

#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML files - Read and write XML files." '=====================================================================================================================================

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.

        dgvInputLocations.AllowUserToAddRows = False 'Otherwise a blank row is added to the saved InputLocationValues

        dgvOutputLocations.AllowUserToAddRows = False

        dgvConversion.AllowUserToAddRows = False


        'Datum Transformation:
        Dim EntryCoordType As String
        If rbEnterInputEastNorth.Checked Then
            EntryCoordType = "Projected"
        ElseIf rbEnterInputLongLat.Checked Then
            EntryCoordType = "Geographic"
        ElseIf rbEnterInputXYZ.Checked Then
            EntryCoordType = "Cartesian"
        Else
            EntryCoordType = "Geographic"
        End If

        Dim InputEntryCoordType As String
        If rbInputEastNorth.Checked Then
            InputEntryCoordType = "Projected"
        ElseIf rbInputLongLat.Checked Then
            InputEntryCoordType = "Geographic"
        ElseIf rbInputXYZ.Checked Then
            InputEntryCoordType = "Cartesian"
        Else
            InputEntryCoordType = "Geographic"
        End If

        Dim InputDegreeDisplay As String
        If rbInputDecDegrees.Checked Then
            InputDegreeDisplay = "DecimalDegrees"
        ElseIf rbInputDMS.Checked Then
            InputDegreeDisplay = "Deg-Min-Sec"
        Else
            InputDegreeDisplay = "DecimalDegrees"
        End If

        Dim OutputDegreeDisplay As String
        If rbOutputDecDegrees.Checked Then
            OutputDegreeDisplay = "DecimalDegrees"
        ElseIf rbOutputDMS.Checked Then
            OutputDegreeDisplay = "Deg-Min-Sec"
        Else
            OutputDegreeDisplay = "DecimalDegrees"
        End If

        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <!--Other Settings-->
                               <FileName><%= txtFileName.Text %></FileName>
                               <DataName><%= txtDataName.Text %></DataName>
                               <DataDescription><%= txtDescription.Text %></DataDescription>
                               <SelectedTabIndex><%= TabControl1.SelectedIndex %></SelectedTabIndex>
                               <!--Input CRS Settings-->
                               <InputEntryCoordType><%= InputEntryCoordType %></InputEntryCoordType>
                               <SelectedInputTabIndex><%= TabControl2.SelectedIndex %></SelectedInputTabIndex>
                               <InputCrsQuery><%= txtInputCrsQuery.Text %></InputCrsQuery>
                               <InputQueryNameContains><%= txtFindInput.Text %></InputQueryNameContains>
                               <InputCrsCode><%= txtInputCrsCode.Text %></InputCrsCode>
                               <SplitDist1><%= SplitContainer1.SplitterDistance %></SplitDist1>
                               <SplitDist2><%= SplitContainer2.SplitterDistance %></SplitDist2>
                               <SplitDist5><%= SplitContainer5.SplitterDistance %></SplitDist5>
                               <SelInputRowNo><%= SelInputRowNo %></SelInputRowNo>
                               <!--Save Input Location values-->
                               <InputDegreeDisplay><%= InputDegreeDisplay %></InputDegreeDisplay>
                               <InputDegreesDecPlaces><%= txtInputDegreeDecPlaces.Text %></InputDegreesDecPlaces>
                               <InputSecondsDecPlaces><%= txtInputSecondsDecPlaces.Text %></InputSecondsDecPlaces>
                               <InputShowDmsSymbols><%= chkInputDmsSymbols.Checked %></InputShowDmsSymbols>
                               <InputHeightFormat><%= txtInputHeightFormat.Text %></InputHeightFormat>
                               <InputProjectedFormat><%= txtInputProjFormat.Text %></InputProjectedFormat>
                               <InputCartesianFormat><%= txtInputCartFormat.Text %></InputCartesianFormat>
                               <InputLocationValues>
                                   <%= From Row As DataGridViewRow In dgvInputLocations.Rows
                                       Select
                                         <Row>
                                             <Easting><%= Row.Cells(0).Value %></Easting>
                                             <Northing><%= Row.Cells(1).Value %></Northing>
                                             <Longitude><%= Row.Cells(2).Value %></Longitude>
                                             <Latitude><%= Row.Cells(3).Value %></Latitude>
                                             <EllipsoidalHeight><%= Row.Cells(4).Value %></EllipsoidalHeight>
                                             <X><%= Row.Cells(5).Value %></X>
                                             <Y><%= Row.Cells(6).Value %></Y>
                                             <Z><%= Row.Cells(7).Value %></Z>
                                         </Row> %>
                               </InputLocationValues>
                               <!--Output CRS Settings-->
                               <SelectedOutputTabIndex><%= TabControl3.SelectedIndex %></SelectedOutputTabIndex>
                               <OutputCrsQuery><%= txtOutputCrsQuery.Text %></OutputCrsQuery>
                               <OutputQueryNameContains><%= txtFindOutput.Text %></OutputQueryNameContains>
                               <OutputCrsCode><%= txtOutputCrsCode.Text %></OutputCrsCode>
                               <SplitDist3><%= SplitContainer3.SplitterDistance %></SplitDist3>
                               <SplitDist4><%= SplitContainer4.SplitterDistance %></SplitDist4>
                               <SplitDist6><%= SplitContainer6.SplitterDistance %></SplitDist6>
                               <SelOutputRowNo><%= SelOutputRowNo %></SelOutputRowNo>
                               <!--Save Output Location values-->
                               <OutputDegreeDisplay><%= OutputDegreeDisplay %></OutputDegreeDisplay>
                               <OutputDegreesDecPlaces><%= txtOutputDegreeDecPlaces.Text %></OutputDegreesDecPlaces>
                               <OutputSecondsDecPlaces><%= txtOutputSecondsDecPlaces.Text %></OutputSecondsDecPlaces>
                               <OutputShowDmsSymbols><%= chkOutputDmsSymbols.Checked %></OutputShowDmsSymbols>
                               <OutputHeightFormat><%= txtOutputHeightFormat.Text %></OutputHeightFormat>
                               <OutputProjectedFormat><%= txtOutputProjFormat.Text %></OutputProjectedFormat>
                               <OutputCartesianFormat><%= txtOutputCartFormat.Text %></OutputCartesianFormat>
                               <OutputLocationValues>
                                   <%= From Row As DataGridViewRow In dgvOutputLocations.Rows
                                       Select
                                         <Row>
                                             <Easting><%= Row.Cells(0).Value %></Easting>
                                             <Northing><%= Row.Cells(1).Value %></Northing>
                                             <Longitude><%= Row.Cells(2).Value %></Longitude>
                                             <Latitude><%= Row.Cells(3).Value %></Latitude>
                                             <EllipsoidalHeight><%= Row.Cells(4).Value %></EllipsoidalHeight>
                                             <X><%= Row.Cells(5).Value %></X>
                                             <Y><%= Row.Cells(6).Value %></Y>
                                             <Z><%= Row.Cells(7).Value %></Z>
                                         </Row> %>
                               </OutputLocationValues>
                               <!--Datum Transformation Settings-->
                               <SelectedDatumTransTabIndex><%= TabControl4.SelectedIndex %></SelectedDatumTransTabIndex>
                               <SelectedViaWgs84TabIndex><%= TabControl5.SelectedIndex %></SelectedViaWgs84TabIndex>
                               <SplitDist7><%= SplitContainer7.SplitterDistance %></SplitDist7>
                               <SplitDist8><%= SplitContainer8.SplitterDistance %></SplitDist8>
                               <SplitDist9><%= SplitContainer9.SplitterDistance %></SplitDist9>
                               <SplitDist10><%= SplitContainer10.SplitterDistance %></SplitDist10>
                               <SplitDist11><%= SplitContainer11.SplitterDistance %></SplitDist11>
                               <SplitDist12><%= SplitContainer12.SplitterDistance %></SplitDist12>
                               <SplitDist13><%= SplitContainer13.SplitterDistance %></SplitDist13>
                               <SplitDist14><%= SplitContainer14.SplitterDistance %></SplitDist14>
                               <SplitDist15><%= SplitContainer15.SplitterDistance %></SplitDist15>
                               <DirectCoordOpCode><%= Conversion.DatumTrans.DirectCoordOp.Code %></DirectCoordOpCode>
                               <InputToWgs84CoordOpCode><%= Conversion.DatumTrans.InputToWgs84CoordOp.Code %></InputToWgs84CoordOpCode>
                               <Wgs84ToOutputCoordOpCode><%= Conversion.DatumTrans.Wgs84ToOutputCoordOp.Code %></Wgs84ToOutputCoordOpCode>
                               <!--  Coordinate Type Conversion Settings-->
                               <DatumTransType><%= Conversion.DatumTrans.Type %></DatumTransType>
                               <!--  Datum Transformation Settings-->
                               <EntryCoordType><%= EntryCoordType %></EntryCoordType>
                               <ShowInputEastingNorthing><%= chkShowInputEastNorth.Checked %></ShowInputEastingNorthing>
                               <ShowInputLongitudeLatitude><%= chkShowInputLongLat.Checked %></ShowInputLongitudeLatitude>
                               <ShowInputXYS><%= chkShowInputXYZ.Checked %></ShowInputXYS>
                               <ShowWgs84XYZ><%= chkShowWgs84XYZ.Checked %></ShowWgs84XYZ>
                               <ShowOutputEastingNorthing><%= chkShowOutputEastNorth.Checked %></ShowOutputEastingNorthing>
                               <ShowOutputLongitudeLatitude><%= chkShowOutputLongLat.Checked %></ShowOutputLongitudeLatitude>
                               <ShowOutputXYZ><%= chkShowOutputXYZ.Checked %></ShowOutputXYZ>
                               <ShowPointNumber><%= chkShowPointNumber.Checked %></ShowPointNumber>
                               <ShowPointName><%= chkShowPointName.Checked %></ShowPointName>
                               <ShowPointDescription><%= chkShowPointDescription.Checked %></ShowPointDescription>
                               <ProjectedFormat><%= txtProjFormat.Text %></ProjectedFormat>
                               <CartesianFormat><%= txtCartFormat.Text %></CartesianFormat>
                               <ShowDegMinSec><%= rbDMS.Checked %></ShowDegMinSec>
                               <ShowDegMinSecSymbols><%= chkDmsSymbols.Checked %></ShowDegMinSecSymbols>
                               <DecDegreesDecPlaces><%= txtDegreeDecPlaces.Text %></DecDegreesDecPlaces>
                               <DmsSecondsDecPlaces><%= txtSecondsDecPlaces.Text %></DmsSecondsDecPlaces>
                               <HeightFormat><%= txtHeightFormat.Text %></HeightFormat>
                               <!--Save the Datum Transformation Input Points-->
                               <%= DatumTransInputData(EntryCoordType).<DatumTransInputData> %>
                           </FormSettings>

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & "_" & FormNo & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    'Private Function DatumTransInputPoints_OLD(EntryCoordType As String) As XDocument
    '    'Returns a XDocument containing the Datum Transformation Input Points shown in dgvConversion.

    '    If EntryCoordType = "Projected" Then
    '        Dim EastingCol As Integer
    '        Dim NorthingCol As Integer
    '        For Each Col As DataGridViewColumn In dgvConversion.Columns
    '            If Col.HeaderText = "Input Easting" Then EastingCol = Col.Index
    '            If Col.HeaderText = "Input Northing" Then NorthingCol = Col.Index
    '        Next
    '        Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
    '                   <DatumTransInputPoints>
    '                       <%= From Row As DataGridViewRow In dgvConversion.Rows
    '                           Select
    '                          <Row>
    '                              <Easting><%= Row.Cells(EastingCol).Value %></Easting>
    '                              <Northing><%= Row.Cells(NorthingCol).Value %></Northing>
    '                          </Row> %>
    '                   </DatumTransInputPoints>
    '        Return XDoc

    '    ElseIf EntryCoordType = "Geographic" Then
    '        Dim LongitudeCol As Integer
    '        Dim LatitudeCol As Integer
    '        Dim HeightCol As Integer
    '        For Each Col As DataGridViewColumn In dgvConversion.Columns
    '            If Col.HeaderText = "Input Longitude" Then LongitudeCol = Col.Index
    '            If Col.HeaderText = "Input Latitude" Then LatitudeCol = Col.Index
    '            If Col.HeaderText = "Input Ellipsoidal Height" Then HeightCol = Col.Index
    '        Next
    '        Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
    '                   <DatumTransInputPoints>
    '                       <%= From Row As DataGridViewRow In dgvConversion.Rows
    '                           Select
    '                          <Row>
    '                              <Longitude><%= Row.Cells(LongitudeCol).Value %></Longitude>
    '                              <Latitude><%= Row.Cells(LatitudeCol).Value %></Latitude>
    '                              <Height><%= Row.Cells(HeightCol).Value %></Height>
    '                          </Row> %>
    '                   </DatumTransInputPoints>
    '        Return XDoc

    '    ElseIf EntryCoordType = "Cartesian" Then
    '        Dim XCol As Integer
    '        Dim YCol As Integer
    '        Dim ZCol As Integer
    '        For Each Col As DataGridViewColumn In dgvConversion.Columns
    '            If Col.HeaderText = "Input X" Then XCol = Col.Index
    '            If Col.HeaderText = "Input Y" Then YCol = Col.Index
    '            If Col.HeaderText = "Input Z" Then ZCol = Col.Index
    '        Next
    '        Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
    '                   <DatumTransInputPoints>
    '                       <%= From Row As DataGridViewRow In dgvConversion.Rows
    '                           Select
    '                          <Row>
    '                              <X><%= Row.Cells(XCol).Value %></X>
    '                              <Y><%= Row.Cells(YCol).Value %></Y>
    '                              <Z><%= Row.Cells(ZCol).Value %></Z>
    '                          </Row> %>
    '                   </DatumTransInputPoints>
    '        Return XDoc

    '    Else

    '    End If

    'End Function

    Private Function DatumTransInputData(EntryCoordType As String) As XDocument
        'Returns a XDocument containing the Datum Transformation Input Points shown in dgvConversion.

        If EntryCoordType = "Projected" Then

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <DatumTransInputData>
                           <%= From Row As DataGridViewRow In dgvConversion.Rows
                               Select
                              <Row>
                                  <PointNumber><%= Row.Cells(0).Value %></PointNumber>
                                  <PointName><%= Row.Cells(1).Value %></PointName>
                                  <PointDescription><%= Row.Cells(2).Value %></PointDescription>
                                  <Easting><%= Row.Cells(3).Value %></Easting>
                                  <Northing><%= Row.Cells(4).Value %></Northing>
                              </Row> %>
                       </DatumTransInputData>
            Return XDoc

        ElseIf EntryCoordType = "Geographic" Then

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <DatumTransInputData>
                           <%= From Row As DataGridViewRow In dgvConversion.Rows
                               Select
                              <Row>
                                  <PointNumber><%= Row.Cells(0).Value %></PointNumber>
                                  <PointName><%= Row.Cells(1).Value %></PointName>
                                  <PointDescription><%= Row.Cells(2).Value %></PointDescription>
                                  <Longitude><%= Row.Cells(5).Value %></Longitude>
                                  <Latitude><%= Row.Cells(6).Value %></Latitude>
                                  <Height><%= Row.Cells(7).Value %></Height>
                              </Row> %>
                       </DatumTransInputData>
            Return XDoc

        ElseIf EntryCoordType = "Cartesian" Then

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <DatumTransInputData>
                           <%= From Row As DataGridViewRow In dgvConversion.Rows
                               Select
                              <Row>
                                  <PointNumber><%= Row.Cells(0).Value %></PointNumber>
                                  <PointName><%= Row.Cells(1).Value %></PointName>
                                  <PointDescription><%= Row.Cells(2).Value %></PointDescription>
                                  <X><%= Row.Cells(8).Value %></X>
                                  <Y><%= Row.Cells(9).Value %></Y>
                                  <Z><%= Row.Cells(10).Value %></Z>
                              </Row> %>
                       </DatumTransInputData>
            Return XDoc

        Else

        End If

    End Function

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

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
            If Settings.<FormSettings>.<FileName>.Value <> Nothing Then txtFileName.Text = Settings.<FormSettings>.<FileName>.Value
            If Settings.<FormSettings>.<DataName>.Value <> Nothing Then txtDataName.Text = Settings.<FormSettings>.<DataName>.Value
            If Settings.<FormSettings>.<DataDescription>.Value <> Nothing Then txtDescription.Text = Settings.<FormSettings>.<DataDescription>.Value

            If Settings.<FormSettings>.<SelectedTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedTabIndex>.Value

            Dim InputEntryCoordType As String
            If Settings.<FormSettings>.<InputEntryCoordType>.Value <> Nothing Then
                InputEntryCoordType = Settings.<FormSettings>.<InputEntryCoordType>.Value
                Select Case InputEntryCoordType
                    Case "Projected"
                        rbInputEastNorth.Checked = True
                    Case "Geographic"
                        rbInputLongLat.Checked = True
                    Case "Cartesian"
                        rbInputXYZ.Checked = True
                End Select
            End If

            If Settings.<FormSettings>.<SelectedInputTabIndex>.Value <> Nothing Then TabControl2.SelectedIndex = Settings.<FormSettings>.<SelectedInputTabIndex>.Value
            If Settings.<FormSettings>.<InputCrsQuery>.Value <> Nothing Then
                txtInputCrsQuery.Text = Settings.<FormSettings>.<InputCrsQuery>.Value
                ApplyInputCrsQuery()
            End If

            If Settings.<FormSettings>.<SelectedOutputTabIndex>.Value <> Nothing Then TabControl3.SelectedIndex = Settings.<FormSettings>.<SelectedOutputTabIndex>.Value
            If Settings.<FormSettings>.<OutputCrsQuery>.Value <> Nothing Then
                txtOutputCrsQuery.Text = Settings.<FormSettings>.<OutputCrsQuery>.Value
                ApplyOutputCrsQuery()
            End If

            If Settings.<FormSettings>.<InputQueryNameContains>.Value <> Nothing Then txtFindInput.Text = Settings.<FormSettings>.<InputQueryNameContains>.Value
            If Settings.<FormSettings>.<OutputQueryNameContains>.Value <> Nothing Then txtFindOutput.Text = Settings.<FormSettings>.<OutputQueryNameContains>.Value

            If Settings.<FormSettings>.<InputCrsCode>.Value <> Nothing Then
                Conversion.InputCrs.Code = Settings.<FormSettings>.<InputCrsCode>.Value
                Conversion.InputCrs.GetAllSourceTargetCoordOps()
                txtInputCrsCode.Text = Conversion.InputCrs.Code
            End If
            If Settings.<FormSettings>.<SplitDist1>.Value <> Nothing Then
                SplitDist1 = Settings.<FormSettings>.<SplitDist1>.Value
                SplitContainer1.SplitterDistance = SplitDist1
            End If

            If Settings.<FormSettings>.<SplitDist2>.Value <> Nothing Then
                SplitDist2 = Settings.<FormSettings>.<SplitDist2>.Value
                SplitContainer2.SplitterDistance = SplitDist2
            End If

            If Settings.<FormSettings>.<OutputCrsCode>.Value <> Nothing Then
                Conversion.OutputCrs.Code = Settings.<FormSettings>.<OutputCrsCode>.Value
                Conversion.OutputCrs.GetAllSourceTargetCoordOps()
                txtOutputCrsCode.Text = Conversion.OutputCrs.Code
            End If

            If Settings.<FormSettings>.<SplitDist3>.Value <> Nothing Then
                SplitDist3 = Settings.<FormSettings>.<SplitDist3>.Value
                SplitContainer3.SplitterDistance = SplitDist3
            End If

            If Settings.<FormSettings>.<SplitDist4>.Value <> Nothing Then
                SplitDist4 = Settings.<FormSettings>.<SplitDist4>.Value
                SplitContainer4.SplitterDistance = SplitDist4
            End If


            If Settings.<FormSettings>.<SelInputRowNo>.Value <> Nothing Then
                SelInputRowNo = Settings.<FormSettings>.<SelInputRowNo>.Value
                If dgvInputCrsList.RowCount > SelInputRowNo And SelInputRowNo > -1 Then
                    dgvInputCrsList.ClearSelection()
                    dgvInputCrsList.Rows(SelInputRowNo).Selected = True
                End If
            End If
            If Settings.<FormSettings>.<SelOutputRowNo>.Value <> Nothing Then
                SelOutputRowNo = Settings.<FormSettings>.<SelOutputRowNo>.Value
                If dgvOutputCrsList.RowCount > SelOutputRowNo And SelOutputRowNo > -1 Then
                    dgvOutputCrsList.ClearSelection()
                    dgvOutputCrsList.Rows(SelOutputRowNo).Selected = True
                End If
            End If

            If Settings.<FormSettings>.<InputDegreeDisplay>.Value <> Nothing Then
                Dim InputDegreeDisplay As String = Settings.<FormSettings>.<InputDegreeDisplay>.Value
                If InputDegreeDisplay = "Deg-Min-Sec" Then
                    rbInputDMS.Checked = True
                Else
                    rbInputDecDegrees.Checked = True
                End If
            End If

            If Settings.<FormSettings>.<InputDegreesDecPlaces>.Value <> Nothing Then
                txtInputDegreeDecPlaces.Text = Settings.<FormSettings>.<InputDegreesDecPlaces>.Value
                dgvInputLocations.Columns(2).DefaultCellStyle.Format = "F" & txtInputDegreeDecPlaces.Text
            End If

            If Settings.<FormSettings>.<InputSecondsDecPlaces>.Value <> Nothing Then
                txtInputSecondsDecPlaces.Text = Settings.<FormSettings>.<InputSecondsDecPlaces>.Value
                Conversion.InputCrs.Coord.DegMinSecDecimalPlaces = txtInputSecondsDecPlaces.Text
            End If

            If Settings.<FormSettings>.<InputShowDmsSymbols>.Value <> Nothing Then
                chkInputDmsSymbols.Checked = Settings.<FormSettings>.<InputShowDmsSymbols>.Value
            End If

            If Settings.<FormSettings>.<InputHeightFormat>.Value <> Nothing Then
                txtInputHeightFormat.Text = Settings.<FormSettings>.<InputHeightFormat>.Value
                dgvInputLocations.Columns(4).DefaultCellStyle.Format = txtInputHeightFormat.Text
            End If

            If Settings.<FormSettings>.<InputProjectedFormat>.Value <> Nothing Then
                txtInputProjFormat.Text = Settings.<FormSettings>.<InputProjectedFormat>.Value
                dgvInputLocations.Columns(0).DefaultCellStyle.Format = txtInputProjFormat.Text
                dgvInputLocations.Columns(1).DefaultCellStyle.Format = txtInputProjFormat.Text
            End If

            If Settings.<FormSettings>.<InputCartesianFormat>.Value <> Nothing Then
                txtInputCartFormat.Text = Settings.<FormSettings>.<InputCartesianFormat>.Value
                dgvInputLocations.Columns(5).DefaultCellStyle.Format = txtInputCartFormat.Text
                dgvInputLocations.Columns(6).DefaultCellStyle.Format = txtInputCartFormat.Text
                dgvInputLocations.Columns(7).DefaultCellStyle.Format = txtInputCartFormat.Text
            End If

            If Settings.<FormSettings>.<InputLocationValues>.Value <> Nothing Then
                Dim InputLocationValues = From Item In Settings.<FormSettings>.<InputLocationValues>.<Row>
                For Each Item In InputLocationValues
                    dgvInputLocations.Rows.Add(Val(Item.<Easting>.Value.Replace(",", "")), Val(Item.<Northing>.Value.Replace(",", "")), Item.<Longitude>.Value, Item.<Latitude>.Value, Item.<EllipsoidalHeight>.Value, Val(Item.<X>.Value.Replace(",", "")), Val(Item.<Y>.Value.Replace(",", "")), Val(Item.<Z>.Value.Replace(",", "")))
                Next
            End If
            dgvInputLocations.AutoResizeColumns()

            If Settings.<FormSettings>.<OutputDegreeDisplay>.Value <> Nothing Then
                Dim OutputDegreeDisplay As String = Settings.<FormSettings>.<OutputDegreeDisplay>.Value
                If OutputDegreeDisplay = "Deg-Min-Sec" Then
                    rbOutputDMS.Checked = True
                Else
                    rbOutputDecDegrees.Checked = True
                End If
            End If

            If Settings.<FormSettings>.<OutputDegreesDecPlaces>.Value <> Nothing Then
                txtOutputDegreeDecPlaces.Text = Settings.<FormSettings>.<OutputDegreesDecPlaces>.Value
                dgvOutputLocations.Columns(2).DefaultCellStyle.Format = "F" & txtOutputDegreeDecPlaces.Text
            End If

            If Settings.<FormSettings>.<OutputSecondsDecPlaces>.Value <> Nothing Then
                txtOutputSecondsDecPlaces.Text = Settings.<FormSettings>.<OutputSecondsDecPlaces>.Value
                Conversion.OutputCrs.Coord.DegMinSecDecimalPlaces = txtOutputSecondsDecPlaces.Text
            End If

            If Settings.<FormSettings>.<OutputShowDmsSymbols>.Value <> Nothing Then
                chkOutputDmsSymbols.Checked = Settings.<FormSettings>.<OutputShowDmsSymbols>.Value
            End If

            If Settings.<FormSettings>.<OutputHeightFormat>.Value <> Nothing Then
                txtOutputHeightFormat.Text = Settings.<FormSettings>.<OutputHeightFormat>.Value
                dgvOutputLocations.Columns(4).DefaultCellStyle.Format = txtOutputHeightFormat.Text
            End If

            If Settings.<FormSettings>.<OutputProjectedFormat>.Value <> Nothing Then
                txtOutputProjFormat.Text = Settings.<FormSettings>.<OutputProjectedFormat>.Value
                dgvOutputLocations.Columns(0).DefaultCellStyle.Format = txtOutputProjFormat.Text
                dgvOutputLocations.Columns(1).DefaultCellStyle.Format = txtOutputProjFormat.Text
            End If

            If Settings.<FormSettings>.<OutputCartesianFormat>.Value <> Nothing Then
                txtOutputCartFormat.Text = Settings.<FormSettings>.<OutputCartesianFormat>.Value
                dgvOutputLocations.Columns(5).DefaultCellStyle.Format = txtOutputCartFormat.Text
                dgvOutputLocations.Columns(6).DefaultCellStyle.Format = txtOutputCartFormat.Text
                dgvOutputLocations.Columns(7).DefaultCellStyle.Format = txtOutputCartFormat.Text
            End If

            If Settings.<FormSettings>.<OutputLocationValues>.Value <> Nothing Then
                Dim OutputLocationValues = From Item In Settings.<FormSettings>.<OutputLocationValues>.<Row>
                For Each Item In OutputLocationValues
                    dgvOutputLocations.Rows.Add(Val(Item.<Easting>.Value.Replace(",", "")), Val(Item.<Northing>.Value.Replace(",", "")), Item.<Longitude>.Value, Item.<Latitude>.Value, Item.<EllipsoidalHeight>.Value, Val(Item.<X>.Value.Replace(",", "")), Val(Item.<Y>.Value.Replace(",", "")), Val(Item.<Z>.Value.Replace(",", "")))
                Next
            End If
            dgvInputLocations.AutoResizeColumns()

            If Settings.<FormSettings>.<SelectedDatumTransTabIndex>.Value <> Nothing Then TabControl4.SelectedIndex = Settings.<FormSettings>.<SelectedDatumTransTabIndex>.Value
            If Settings.<FormSettings>.<SelectedViaWgs84TabIndex>.Value <> Nothing Then TabControl5.SelectedIndex = Settings.<FormSettings>.<SelectedViaWgs84TabIndex>.Value
            If Settings.<FormSettings>.<SplitDist5>.Value <> Nothing Then
                SplitDist5 = Settings.<FormSettings>.<SplitDist5>.Value
                SplitContainer5.SplitterDistance = SplitDist5
            End If
            If Settings.<FormSettings>.<SplitDist6>.Value <> Nothing Then
                SplitDist6 = Settings.<FormSettings>.<SplitDist6>.Value
                SplitContainer6.SplitterDistance = SplitDist6
            End If

            If Settings.<FormSettings>.<SplitDist7>.Value <> Nothing Then
                SplitDist7 = Settings.<FormSettings>.<SplitDist7>.Value
                SplitContainer7.SplitterDistance = SplitDist7
            End If

            If Settings.<FormSettings>.<SplitDist8>.Value <> Nothing Then
                SplitDist8 = Settings.<FormSettings>.<SplitDist8>.Value
                SplitContainer8.SplitterDistance = SplitDist8
            End If
            If Settings.<FormSettings>.<SplitDist9>.Value <> Nothing Then
                SplitDist9 = Settings.<FormSettings>.<SplitDist9>.Value
                SplitContainer9.SplitterDistance = SplitDist9
            End If
            If Settings.<FormSettings>.<SplitDist10>.Value <> Nothing Then
                SplitDist10 = Settings.<FormSettings>.<SplitDist10>.Value
                SplitContainer10.SplitterDistance = SplitDist10
            End If
            If Settings.<FormSettings>.<SplitDist11>.Value <> Nothing Then
                SplitDist11 = Settings.<FormSettings>.<SplitDist11>.Value
                SplitContainer11.SplitterDistance = SplitDist11
            End If

            If Settings.<FormSettings>.<SplitDist12>.Value <> Nothing Then
                SplitDist12 = Settings.<FormSettings>.<SplitDist12>.Value
                SplitContainer12.SplitterDistance = SplitDist12
            End If

            If Settings.<FormSettings>.<SplitDist13>.Value <> Nothing Then
                SplitDist13 = Settings.<FormSettings>.<SplitDist13>.Value
                SplitContainer13.SplitterDistance = SplitDist13
            End If

            If Settings.<FormSettings>.<SplitDist14>.Value <> Nothing Then
                SplitDist14 = Settings.<FormSettings>.<SplitDist14>.Value
                SplitContainer14.SplitterDistance = SplitDist14
            End If
            If Settings.<FormSettings>.<SplitDist15>.Value <> Nothing Then
                SplitDist15 = Settings.<FormSettings>.<SplitDist15>.Value
                SplitContainer15.SplitterDistance = SplitDist15
            End If
            If Settings.<FormSettings>.<DirectCoordOpCode>.Value <> Nothing Then DirectCoordOpCode = Settings.<FormSettings>.<DirectCoordOpCode>.Value
            If Settings.<FormSettings>.<InputToWgs84CoordOpCode>.Value <> Nothing Then InputToWgs84CoordOpCode = Settings.<FormSettings>.<InputToWgs84CoordOpCode>.Value
            If Settings.<FormSettings>.<Wgs84ToOutputCoordOpCode>.Value <> Nothing Then Wgs84ToOutputCoordOpCode = Settings.<FormSettings>.<Wgs84ToOutputCoordOpCode>.Value
            Dim EntryCoordType As String
            If Settings.<FormSettings>.<EntryCoordType>.Value <> Nothing Then
                EntryCoordType = Settings.<FormSettings>.<EntryCoordType>.Value
                'Select Case Settings.<FormSettings>.<EntryCoordType>.Value
                Select Case EntryCoordType
                    Case "Projected"
                        rbEnterInputEastNorth.Checked = True
                    Case "Geographic"
                        rbEnterInputLongLat.Checked = True
                    Case "Cartesian"
                        rbEnterInputXYZ.Checked = True
                End Select
            End If
            If Settings.<FormSettings>.<ShowPointNumber>.Value <> Nothing Then chkShowPointNumber.Checked = Settings.<FormSettings>.<ShowPointNumber>.Value
            If Settings.<FormSettings>.<ShowPointName>.Value <> Nothing Then chkShowPointName.Checked = Settings.<FormSettings>.<ShowPointName>.Value
            If Settings.<FormSettings>.<ShowPointDescription>.Value <> Nothing Then chkShowPointDescription.Checked = Settings.<FormSettings>.<ShowPointDescription>.Value
            If Settings.<FormSettings>.<ShowInputEastingNorthing>.Value <> Nothing Then chkShowInputEastNorth.Checked = Settings.<FormSettings>.<ShowInputEastingNorthing>.Value
            If Settings.<FormSettings>.<ShowInputLongitudeLatitude>.Value <> Nothing Then chkShowInputLongLat.Checked = Settings.<FormSettings>.<ShowInputLongitudeLatitude>.Value
            If Settings.<FormSettings>.<ShowInputXYS>.Value <> Nothing Then chkShowInputXYZ.Checked = Settings.<FormSettings>.<ShowInputXYS>.Value
            If Settings.<FormSettings>.<ShowWgs84XYZ>.Value <> Nothing Then chkShowWgs84XYZ.Checked = Settings.<FormSettings>.<ShowWgs84XYZ>.Value
            If Settings.<FormSettings>.<ShowOutputEastingNorthing>.Value <> Nothing Then chkShowOutputEastNorth.Checked = Settings.<FormSettings>.<ShowOutputEastingNorthing>.Value
            If Settings.<FormSettings>.<ShowOutputLongitudeLatitude>.Value <> Nothing Then chkShowOutputLongLat.Checked = Settings.<FormSettings>.<ShowOutputLongitudeLatitude>.Value
            If Settings.<FormSettings>.<ShowOutputXYZ>.Value <> Nothing Then chkShowOutputXYZ.Checked = Settings.<FormSettings>.<ShowOutputXYZ>.Value
            If Settings.<FormSettings>.<ProjectedFormat>.Value <> Nothing Then txtProjFormat.Text = Settings.<FormSettings>.<ProjectedFormat>.Value
            If Settings.<FormSettings>.<CartesianFormat>.Value <> Nothing Then txtCartFormat.Text = Settings.<FormSettings>.<CartesianFormat>.Value
            'If Settings.<FormSettings>.<DecDegreesDecPlaces>.Value <> Nothing Then txtDegreeDecPlaces.Text = Settings.<FormSettings>.<DecDegreesDecPlaces>.Value
            If Settings.<FormSettings>.<DecDegreesDecPlaces>.Value <> Nothing Then
                txtDegreeDecPlaces.Text = Settings.<FormSettings>.<DecDegreesDecPlaces>.Value
                If rbDecDegrees.Checked Then
                    Try
                        txtDegreeDecPlaces.Text = Int(txtDegreeDecPlaces.Text.Trim)
                        Dim Format As String = "F" & txtDegreeDecPlaces.Text.Trim
                        dgvConversion.Columns(5).DefaultCellStyle.Format = Format
                        dgvConversion.Columns(6).DefaultCellStyle.Format = Format
                        dgvConversion.Columns(17).DefaultCellStyle.Format = Format
                        dgvConversion.Columns(18).DefaultCellStyle.Format = Format
                    Catch ex As Exception

                    End Try
                End If
            End If
            If Settings.<FormSettings>.<DmsSecondsDecPlaces>.Value <> Nothing Then
                txtSecondsDecPlaces.Text = Settings.<FormSettings>.<DmsSecondsDecPlaces>.Value
                Conversion.InputCrs.Coord.DegMinSecDecimalPlaces = txtSecondsDecPlaces.Text
                Conversion.OutputCrs.Coord.DegMinSecDecimalPlaces = txtSecondsDecPlaces.Text
            End If
            If Settings.<FormSettings>.<HeightFormat>.Value <> Nothing Then
                txtHeightFormat.Text = Settings.<FormSettings>.<HeightFormat>.Value
            End If

            If Settings.<FormSettings>.<ShowDegMinSec>.Value <> Nothing Then rbDMS.Checked = Settings.<FormSettings>.<ShowDegMinSec>.Value
            If Settings.<FormSettings>.<ShowDegMinSecSymbols>.Value <> Nothing Then chkDmsSymbols.Checked = Settings.<FormSettings>.<ShowDegMinSecSymbols>.Value
            If EntryCoordType = "Projected" Then
                Dim RowNo As Integer
                If Settings.<FormSettings>.<DatumTransInputData>.Value <> Nothing Then
                    Dim InputPoints = From Item In Settings.<FormSettings>.<DatumTransInputData>.<Row>
                    For Each Item In InputPoints
                        RowNo = dgvConversion.Rows.Add()
                        If Item.<PointNumber>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(0).Value = Item.<PointNumber>.Value
                        If Item.<PointName>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(1).Value = Item.<PointName>.Value
                        If Item.<PointDescription>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(2).Value = Item.<PointDescription>.Value
                        dgvConversion.Rows(RowNo).Cells(3).Value = Item.<Easting>.Value
                        dgvConversion.Rows(RowNo).Cells(4).Value = Item.<Northing>.Value
                    Next
                End If

            ElseIf EntryCoordType = "Geographic" Then

                Dim RowNo As Integer

                If Settings.<FormSettings>.<DatumTransInputData>.Value <> Nothing Then
                    Dim InputPoints = From Item In Settings.<FormSettings>.<DatumTransInputData>.<Row>
                    For Each Item In InputPoints
                        RowNo = dgvConversion.Rows.Add()
                        If Item.<PointNumber>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(0).Value = Item.<PointNumber>.Value
                        If Item.<PointName>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(1).Value = Item.<PointName>.Value
                        If Item.<PointDescription>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(2).Value = Item.<PointDescription>.Value
                        dgvConversion.Rows(RowNo).Cells(5).Value = Item.<Longitude>.Value
                        dgvConversion.Rows(RowNo).Cells(6).Value = Item.<Latitude>.Value
                        dgvConversion.Rows(RowNo).Cells(7).Value = Item.<Height>.Value
                    Next
                End If

            ElseIf EntryCoordType = "Cartesian" Then
                Dim RowNo As Integer
                If Settings.<FormSettings>.<DatumTransInputData>.Value <> Nothing Then
                    Dim InputPoints = From Item In Settings.<FormSettings>.<DatumTransInputData>.<Row>
                    For Each Item In InputPoints
                        RowNo = dgvConversion.Rows.Add()
                        If Item.<PointNumber>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(0).Value = Item.<PointNumber>.Value
                        If Item.<PointName>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(1).Value = Item.<PointName>.Value
                        If Item.<PointDescription>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(2).Value = Item.<PointDescription>.Value
                        dgvConversion.Rows(RowNo).Cells(8).Value = Item.<X>.Value
                        dgvConversion.Rows(RowNo).Cells(9).Value = Item.<Y>.Value
                        dgvConversion.Rows(RowNo).Cells(10).Value = Item.<Z>.Value
                    Next
                End If

            End If
            ApplyDatumTransFormats()

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

    'Protected Overrides Sub WndProc(ByRef m As Message) 'Save the form settings before the form is minimised:
    '    If m.Msg = &H112 Then 'SysCommand
    '        If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
    '            SaveFormSettings()
    '        End If
    '    End If
    '    MyBase.WndProc(m)
    'End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message) 'Save the form settings before the form is minimised:
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

        udInputRowNo.Minimum = -1
        udInputRowNo.Increment = 1

        udInputDatumLevel.Minimum = 0
        udInputDatumLevel.Maximum = 0
        txtInputDatumNLevels.Text = "0"
        udInputDatumLevel.Increment = 1
        udInputDatumLevel.Value = 0

        udOutputRowNo.Minimum = -1
        udOutputRowNo.Increment = 1

        udOutputDatumLevel.Minimum = 0
        udOutputDatumLevel.Maximum = 0
        txtOutputDatumNLevels.Text = "0"
        udOutputDatumLevel.Increment = 1
        udOutputDatumLevel.Value = 0

        Dim NewFont As New Font("Arial", 11, FontStyle.Regular)

        dgvInputLocations.ColumnCount = 8
        dgvInputLocations.Columns(0).HeaderText = "Easting"
        dgvInputLocations.Columns(0).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(0).Width = 120
        dgvInputLocations.Columns(0).ReadOnly = True
        dgvInputLocations.Columns(0).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(0).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvInputLocations.Columns(1).HeaderText = "Northing"
        dgvInputLocations.Columns(1).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(1).Width = 120
        dgvInputLocations.Columns(1).ReadOnly = True
        dgvInputLocations.Columns(1).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(1).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvInputLocations.Columns(2).HeaderText = "Longitude"
        dgvInputLocations.Columns(2).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(2).Width = 120
        dgvInputLocations.Columns(2).ReadOnly = True
        dgvInputLocations.Columns(2).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(2).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvInputLocations.Columns(3).HeaderText = "Latitude"
        dgvInputLocations.Columns(3).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(3).Width = 120
        dgvInputLocations.Columns(3).ReadOnly = True
        dgvInputLocations.Columns(3).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(3).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvInputLocations.Columns(4).HeaderText = "Ellipsoidal Height"
        dgvInputLocations.Columns(4).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(4).Width = 120
        dgvInputLocations.Columns(4).ReadOnly = True
        dgvInputLocations.Columns(4).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(4).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvInputLocations.Columns(5).HeaderText = "X"
        dgvInputLocations.Columns(5).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(5).Width = 120
        dgvInputLocations.Columns(5).ReadOnly = True
        dgvInputLocations.Columns(5).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(5).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvInputLocations.Columns(6).HeaderText = "Y"
        dgvInputLocations.Columns(6).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(6).Width = 120
        dgvInputLocations.Columns(6).ReadOnly = True
        dgvInputLocations.Columns(6).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(6).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvInputLocations.Columns(7).HeaderText = "Z"
        dgvInputLocations.Columns(7).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(7).Width = 120
        dgvInputLocations.Columns(7).ReadOnly = True
        dgvInputLocations.Columns(7).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(7).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvInputLocations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvInputLocations.AutoResizeColumns()

        dgvOutputLocations.ColumnCount = 8
        dgvOutputLocations.Columns(0).HeaderText = "Easting"
        dgvOutputLocations.Columns(0).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(0).Width = 120
        dgvOutputLocations.Columns(0).ReadOnly = True
        dgvOutputLocations.Columns(0).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(0).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvOutputLocations.Columns(1).HeaderText = "Northing"
        dgvOutputLocations.Columns(1).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(1).Width = 120
        dgvOutputLocations.Columns(1).ReadOnly = True
        dgvOutputLocations.Columns(1).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(1).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvOutputLocations.Columns(2).HeaderText = "Longitude"
        dgvOutputLocations.Columns(2).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(2).Width = 120
        dgvOutputLocations.Columns(2).ReadOnly = True
        dgvOutputLocations.Columns(2).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(2).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvOutputLocations.Columns(3).HeaderText = "Latitude"
        dgvOutputLocations.Columns(3).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(3).Width = 120
        dgvOutputLocations.Columns(3).ReadOnly = True
        dgvOutputLocations.Columns(3).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(3).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvOutputLocations.Columns(4).HeaderText = "Ellipsoidal Height"
        dgvOutputLocations.Columns(4).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(4).Width = 120
        dgvOutputLocations.Columns(4).ReadOnly = True
        dgvOutputLocations.Columns(4).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(4).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvOutputLocations.Columns(5).HeaderText = "X"
        dgvOutputLocations.Columns(5).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(5).Width = 120
        dgvOutputLocations.Columns(5).ReadOnly = True
        dgvOutputLocations.Columns(5).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(5).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvOutputLocations.Columns(6).HeaderText = "Y"
        dgvOutputLocations.Columns(6).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(6).Width = 120
        dgvOutputLocations.Columns(6).ReadOnly = True
        dgvOutputLocations.Columns(6).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(6).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvOutputLocations.Columns(7).HeaderText = "Z"
        dgvOutputLocations.Columns(7).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(7).Width = 120
        dgvOutputLocations.Columns(7).ReadOnly = True
        dgvOutputLocations.Columns(7).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(7).DefaultCellStyle.BackColor = Color.WhiteSmoke
        dgvOutputLocations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvOutputLocations.AutoResizeColumns()

        dgvConversion.ColumnCount = 22
        dgvConversion.Columns(0).HeaderText = "Point Number"
        dgvConversion.Columns(0).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(0).ReadOnly = True
        dgvConversion.Columns(0).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(0).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(1).HeaderText = "Point Name"
        dgvConversion.Columns(1).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(1).ReadOnly = True
        dgvConversion.Columns(1).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(1).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(2).HeaderText = "Point Description"
        dgvConversion.Columns(2).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(2).ReadOnly = True
        dgvConversion.Columns(2).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(2).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(3).HeaderText = "Input Easting"
        dgvConversion.Columns(3).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(3).ReadOnly = True
        dgvConversion.Columns(3).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(3).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(4).HeaderText = "Input Northing"
        dgvConversion.Columns(4).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(4).ReadOnly = True
        dgvConversion.Columns(4).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(4).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(5).HeaderText = "Input Longitude"
        dgvConversion.Columns(5).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(5).ReadOnly = True
        dgvConversion.Columns(5).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(5).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(6).HeaderText = "Input Latitude"
        dgvConversion.Columns(6).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(6).ReadOnly = True
        dgvConversion.Columns(6).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(6).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(7).HeaderText = "Input Ellipsoidal Height"
        dgvConversion.Columns(7).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(7).ReadOnly = True
        dgvConversion.Columns(7).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(7).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(8).HeaderText = "Input X"
        dgvConversion.Columns(8).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(8).ReadOnly = True
        dgvConversion.Columns(8).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(8).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(9).HeaderText = "Input Y"
        dgvConversion.Columns(9).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(9).ReadOnly = True
        dgvConversion.Columns(9).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(9).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(10).HeaderText = "Input Z"
        dgvConversion.Columns(10).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(10).ReadOnly = True
        dgvConversion.Columns(10).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(10).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(11).HeaderText = "WGS 84 X"
        dgvConversion.Columns(11).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(1).ReadOnly = True
        dgvConversion.Columns(1).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(11).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(12).HeaderText = "WGS 84 Y"
        dgvConversion.Columns(12).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(12).ReadOnly = True
        dgvConversion.Columns(12).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(12).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(13).HeaderText = "WGS 84 Z"
        dgvConversion.Columns(13).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(13).ReadOnly = True
        dgvConversion.Columns(13).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(13).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(14).HeaderText = "Output X"
        dgvConversion.Columns(14).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(14).ReadOnly = True
        dgvConversion.Columns(14).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(14).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(15).HeaderText = "Output Y"
        dgvConversion.Columns(15).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(15).ReadOnly = True
        dgvConversion.Columns(15).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(15).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(16).HeaderText = "Output Z"
        dgvConversion.Columns(16).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(16).ReadOnly = True
        dgvConversion.Columns(16).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(16).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(17).HeaderText = "Output Longitude"
        dgvConversion.Columns(17).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(17).ReadOnly = True
        dgvConversion.Columns(17).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(17).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(18).HeaderText = "Output Latitude"
        dgvConversion.Columns(18).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(18).ReadOnly = True
        dgvConversion.Columns(18).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(18).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(19).HeaderText = "Output Ellipsoidal Height"
        dgvConversion.Columns(19).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(19).ReadOnly = True
        dgvConversion.Columns(19).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(19).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(20).HeaderText = "Output Easting"
        dgvConversion.Columns(20).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(20).ReadOnly = True
        dgvConversion.Columns(20).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(20).DefaultCellStyle.BackColor = Color.White
        dgvConversion.Columns(21).HeaderText = "Output Northing"
        dgvConversion.Columns(21).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(21).ReadOnly = True
        dgvConversion.Columns(21).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(21).DefaultCellStyle.BackColor = Color.White
        dgvConversion.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvConversion.AutoResizeColumns()

        udPointFontSize.Minimum = 8
        udPointFontSize.Maximum = 24
        udPointFontSize.Increment = 0.5
        udPointFontSize.DecimalPlaces = 1
        udPointFontSize.Value = 11

        rbDirectDatumTrans.Checked = True 'Select the Direct Datum Transformation button.
        TabControl4.SelectedIndex = 0 'Select the Direct Datum Transformation tab.

        rbEnterInputLongLat.Checked = True
        chkShowInputLongLat.Checked = True
        chkShowOutputLongLat.Checked = True

        rbDecDegrees.Checked = True

        rbInputLongLat.Checked = True
        rbInputDecDegrees.Checked = True
        rbOutputLongLat.Checked = True
        rbOutputDecDegrees.Checked = True

        RestoreFormSettings()   'Restore the form settings

        UpdateDatumTransTable()



        rbInputCrsCoordSysAx1Name.Checked = True

        udInputFont.Minimum = 8
        udInputFont.Maximum = 24
        udInputFont.Increment = 0.5
        udInputFont.DecimalPlaces = 1
        udInputFont.Value = 10

        udInputCalcFontSize.Minimum = 8
        udInputCalcFontSize.Maximum = 24
        udInputCalcFontSize.Increment = 0.5
        udInputCalcFontSize.DecimalPlaces = 1
        udInputCalcFontSize.Value = 11

        dgvInputCrsProjParams.ColumnCount = 7
        dgvInputCrsProjParams.Columns(0).HeaderText = "Code"
        dgvInputCrsProjParams.Columns(1).HeaderText = "Sort Order"
        dgvInputCrsProjParams.Columns(2).HeaderText = "Name"
        dgvInputCrsProjParams.Columns(3).HeaderText = "Sign Reversal"
        dgvInputCrsProjParams.Columns(4).HeaderText = "Value"
        dgvInputCrsProjParams.Columns(5).HeaderText = "Units"
        dgvInputCrsProjParams.Columns(6).HeaderText = "Description"
        dgvInputCrsProjParams.Columns(6).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvInputCrsProjParams.AutoResizeColumns()

        If CrsInputSearch.Tables.Contains("List") Then
            udInputRowNo.Maximum = CrsInputSearch.Tables("List").Rows.Count - 1
            udInputRowNo.Value = SelInputRowNo
        Else
            udInputRowNo.Maximum = -1
            udInputRowNo.Value = -1
        End If

        dgvInputSourceCoordOps.ColumnCount = 10
        dgvInputSourceCoordOps.Columns(0).HeaderText = "Level"
        dgvInputSourceCoordOps.Columns(1).HeaderText = "Name"
        dgvInputSourceCoordOps.Columns(2).HeaderText = "Type"
        dgvInputSourceCoordOps.Columns(3).HeaderText = "Code"
        dgvInputSourceCoordOps.Columns(4).HeaderText = "Accuracy"
        dgvInputSourceCoordOps.Columns(5).HeaderText = "Deprecated"
        dgvInputSourceCoordOps.Columns(6).HeaderText = "Version"
        dgvInputSourceCoordOps.Columns(7).HeaderText = "Revision Date"
        dgvInputSourceCoordOps.Columns(7).DefaultCellStyle.Format = "dd/MM/yyyy"
        dgvInputSourceCoordOps.Columns(8).HeaderText = "Source CRS Code"
        dgvInputSourceCoordOps.Columns(9).HeaderText = "Target CRS Code"
        dgvInputSourceCoordOps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvInputSourceCoordOps.AutoResizeColumns()

        dgvInputTargetCoordOps.ColumnCount = 10
        dgvInputTargetCoordOps.Columns(0).HeaderText = "Level"
        dgvInputTargetCoordOps.Columns(1).HeaderText = "Name"
        dgvInputTargetCoordOps.Columns(2).HeaderText = "Type"
        dgvInputTargetCoordOps.Columns(3).HeaderText = "Code"
        dgvInputTargetCoordOps.Columns(4).HeaderText = "Accuracy"
        dgvInputTargetCoordOps.Columns(5).HeaderText = "Deprecated"
        dgvInputTargetCoordOps.Columns(6).HeaderText = "Version"
        dgvInputTargetCoordOps.Columns(7).HeaderText = "Revision Date"
        dgvInputTargetCoordOps.Columns(7).DefaultCellStyle.Format = "dd/MM/yyyy"
        dgvInputTargetCoordOps.Columns(8).HeaderText = "Source CRS Code"
        dgvInputTargetCoordOps.Columns(9).HeaderText = "Target CRS Code"
        dgvInputTargetCoordOps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvInputTargetCoordOps.AutoResizeColumns()

        rbOutputCrsCoordSysAx1Name.Checked = True

        udOutputFont.Minimum = 8
        udOutputFont.Maximum = 24
        udOutputFont.Increment = 0.5
        udOutputFont.DecimalPlaces = 1
        udOutputFont.Value = 10

        udOutputCalcFontSize.Minimum = 8
        udOutputCalcFontSize.Maximum = 24
        udOutputCalcFontSize.Increment = 0.5
        udOutputCalcFontSize.DecimalPlaces = 1
        udOutputCalcFontSize.Value = 11

        dgvOutputCrsProjParams.ColumnCount = 7
        dgvOutputCrsProjParams.Columns(0).HeaderText = "Code"
        dgvOutputCrsProjParams.Columns(1).HeaderText = "Sort Order"
        dgvOutputCrsProjParams.Columns(2).HeaderText = "Name"
        dgvOutputCrsProjParams.Columns(3).HeaderText = "Sign Reversal"
        dgvOutputCrsProjParams.Columns(4).HeaderText = "Value"
        dgvOutputCrsProjParams.Columns(5).HeaderText = "Units"
        dgvOutputCrsProjParams.Columns(6).HeaderText = "Description"
        dgvOutputCrsProjParams.Columns(6).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvOutputCrsProjParams.AutoResizeColumns()

        If CrsOutputSearch.Tables.Contains("List") Then
            udOutputRowNo.Maximum = CrsOutputSearch.Tables("List").Rows.Count - 1
            udOutputRowNo.Value = SelOutputRowNo
        Else
            udOutputRowNo.Maximum = -1
            udOutputRowNo.Value = -1
        End If

        dgvOutputSourceCoordOps.ColumnCount = 10
        dgvOutputSourceCoordOps.Columns(0).HeaderText = "Level"
        dgvOutputSourceCoordOps.Columns(1).HeaderText = "Name"
        dgvOutputSourceCoordOps.Columns(2).HeaderText = "Type"
        dgvOutputSourceCoordOps.Columns(3).HeaderText = "Code"
        dgvOutputSourceCoordOps.Columns(4).HeaderText = "Accuracy"
        dgvOutputSourceCoordOps.Columns(5).HeaderText = "Deprecated"
        dgvOutputSourceCoordOps.Columns(6).HeaderText = "Version"
        dgvOutputSourceCoordOps.Columns(7).HeaderText = "Revision Date"
        dgvOutputSourceCoordOps.Columns(7).DefaultCellStyle.Format = "dd/MM/yyyy"
        dgvOutputSourceCoordOps.Columns(8).HeaderText = "Source CRS Code"
        dgvOutputSourceCoordOps.Columns(9).HeaderText = "Target CRS Code"
        dgvOutputSourceCoordOps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvOutputSourceCoordOps.AutoResizeColumns()

        dgvOutputTargetCoordOps.ColumnCount = 10
        dgvOutputTargetCoordOps.Columns(0).HeaderText = "Level"
        dgvOutputTargetCoordOps.Columns(1).HeaderText = "Name"
        dgvOutputTargetCoordOps.Columns(2).HeaderText = "Type"
        dgvOutputTargetCoordOps.Columns(3).HeaderText = "Code"
        dgvOutputTargetCoordOps.Columns(4).HeaderText = "Accuracy"
        dgvOutputTargetCoordOps.Columns(5).HeaderText = "Deprecated"
        dgvOutputTargetCoordOps.Columns(6).HeaderText = "Version"
        dgvOutputTargetCoordOps.Columns(7).HeaderText = "Revision Date"
        dgvOutputTargetCoordOps.Columns(7).DefaultCellStyle.Format = "dd/MM/yyyy"
        dgvOutputTargetCoordOps.Columns(8).HeaderText = "Source CRS Code"
        dgvOutputTargetCoordOps.Columns(9).HeaderText = "Target CRS Code"
        dgvOutputTargetCoordOps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvOutputTargetCoordOps.AutoResizeColumns()

        'dgvDirectDTOps.ColumnCount = 13
        dgvDirectDTOps.ColumnCount = 15
        dgvDirectDTOps.Columns(0).HeaderText = "Name"
        dgvDirectDTOps.Columns(1).HeaderText = "Type"
        dgvDirectDTOps.Columns(2).HeaderText = "Code"
        dgvDirectDTOps.Columns(3).HeaderText = "Accuracy"
        dgvDirectDTOps.Columns(4).HeaderText = "Deprecated"
        dgvDirectDTOps.Columns(5).HeaderText = "Version"
        dgvDirectDTOps.Columns(6).HeaderText = "Revision Date"
        dgvDirectDTOps.Columns(6).DefaultCellStyle.Format = "dd/MM/yyyy"
        dgvDirectDTOps.Columns(7).HeaderText = "Source CRS Level"
        dgvDirectDTOps.Columns(8).HeaderText = "Source CRS Code"
        dgvDirectDTOps.Columns(9).HeaderText = "Target CRS Level"
        dgvDirectDTOps.Columns(10).HeaderText = "Target CRS Code"
        dgvDirectDTOps.Columns(11).HeaderText = "Reversible"

        dgvDirectDTOps.Columns(12).HeaderText = "Apply Rev" 'If True the reverse coordinate transformation is applied.

        dgvDirectDTOps.Columns(13).HeaderText = "Method Code"
        dgvDirectDTOps.Columns(14).HeaderText = "Method Name"

        dgvDirectDTOps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvDirectDTOps.AutoResizeColumns()

        'dgvInputToWgs84DTOps.ColumnCount = 13
        dgvInputToWgs84DTOps.ColumnCount = 15
        dgvInputToWgs84DTOps.Columns(0).HeaderText = "Name"
        dgvInputToWgs84DTOps.Columns(1).HeaderText = "Type"
        dgvInputToWgs84DTOps.Columns(2).HeaderText = "Code"
        dgvInputToWgs84DTOps.Columns(3).HeaderText = "Accuracy"
        dgvInputToWgs84DTOps.Columns(4).HeaderText = "Deprecated"
        dgvInputToWgs84DTOps.Columns(5).HeaderText = "Version"
        dgvInputToWgs84DTOps.Columns(6).HeaderText = "Revision Date"
        dgvInputToWgs84DTOps.Columns(6).DefaultCellStyle.Format = "dd/MM/yyyy"
        dgvInputToWgs84DTOps.Columns(7).HeaderText = "Source CRS Level"
        dgvInputToWgs84DTOps.Columns(8).HeaderText = "Source CRS Code"
        dgvInputToWgs84DTOps.Columns(9).HeaderText = "Target CRS Level"
        dgvInputToWgs84DTOps.Columns(10).HeaderText = "Target CRS Code"
        dgvInputToWgs84DTOps.Columns(11).HeaderText = "Reversible"

        dgvInputToWgs84DTOps.Columns(12).HeaderText = "Apply Rev" 'If True the reverse coordinate transformation is applied.

        dgvInputToWgs84DTOps.Columns(13).HeaderText = "Method Code"
        dgvInputToWgs84DTOps.Columns(14).HeaderText = "Method Name"

        dgvInputToWgs84DTOps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvInputToWgs84DTOps.AutoResizeColumns()

        'dgvWgs84ToOutputDTOps.ColumnCount = 13
        dgvWgs84ToOutputDTOps.ColumnCount = 15
        dgvWgs84ToOutputDTOps.Columns(0).HeaderText = "Name"
        dgvWgs84ToOutputDTOps.Columns(1).HeaderText = "Type"
        dgvWgs84ToOutputDTOps.Columns(2).HeaderText = "Code"
        dgvWgs84ToOutputDTOps.Columns(3).HeaderText = "Accuracy"
        dgvWgs84ToOutputDTOps.Columns(4).HeaderText = "Deprecated"
        dgvWgs84ToOutputDTOps.Columns(5).HeaderText = "Version"
        dgvWgs84ToOutputDTOps.Columns(6).HeaderText = "Revision Date"
        dgvWgs84ToOutputDTOps.Columns(6).DefaultCellStyle.Format = "dd/MM/yyyy"
        dgvWgs84ToOutputDTOps.Columns(7).HeaderText = "Source CRS Level"
        dgvWgs84ToOutputDTOps.Columns(8).HeaderText = "Source CRS Code"
        dgvWgs84ToOutputDTOps.Columns(9).HeaderText = "Target CRS Level"
        dgvWgs84ToOutputDTOps.Columns(10).HeaderText = "Target CRS Code"
        dgvWgs84ToOutputDTOps.Columns(11).HeaderText = "Reversible"

        dgvWgs84ToOutputDTOps.Columns(12).HeaderText = "Apply Rev" 'If True the reverse coordinate transformation is applied.

        dgvWgs84ToOutputDTOps.Columns(13).HeaderText = "Method Code"
        dgvWgs84ToOutputDTOps.Columns(14).HeaderText = "Method Name"

        dgvWgs84ToOutputDTOps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvWgs84ToOutputDTOps.AutoResizeColumns()

        dgvDirectDTParams.ColumnCount = 7
        dgvDirectDTParams.Columns(0).HeaderText = "Code"
        dgvDirectDTParams.Columns(1).HeaderText = "Sort Order"
        dgvDirectDTParams.Columns(2).HeaderText = "Name"
        dgvDirectDTParams.Columns(3).HeaderText = "Sign Reversal"
        dgvDirectDTParams.Columns(4).HeaderText = "Value"
        dgvDirectDTParams.Columns(5).HeaderText = "Units"
        dgvDirectDTParams.Columns(6).HeaderText = "Description"
        dgvDirectDTParams.Columns(6).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvDirectDTParams.AutoResizeColumns()

        dgvInputToWgs84DTParams.ColumnCount = 7
        dgvInputToWgs84DTParams.Columns(0).HeaderText = "Code"
        dgvInputToWgs84DTParams.Columns(1).HeaderText = "Sort Order"
        dgvInputToWgs84DTParams.Columns(2).HeaderText = "Name"
        dgvInputToWgs84DTParams.Columns(3).HeaderText = "Sign Reversal"
        dgvInputToWgs84DTParams.Columns(4).HeaderText = "Value"
        dgvInputToWgs84DTParams.Columns(5).HeaderText = "Units"
        dgvInputToWgs84DTParams.Columns(6).HeaderText = "Description"
        dgvInputToWgs84DTParams.Columns(6).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvInputToWgs84DTParams.AutoResizeColumns()

        dgvWgs84ToOutputDTParams.ColumnCount = 7
        dgvWgs84ToOutputDTParams.Columns(0).HeaderText = "Code"
        dgvWgs84ToOutputDTParams.Columns(1).HeaderText = "Sort Order"
        dgvWgs84ToOutputDTParams.Columns(2).HeaderText = "Name"
        dgvWgs84ToOutputDTParams.Columns(3).HeaderText = "Sign Reversal"
        dgvWgs84ToOutputDTParams.Columns(4).HeaderText = "Value"
        dgvWgs84ToOutputDTParams.Columns(5).HeaderText = "Units"
        dgvWgs84ToOutputDTParams.Columns(6).HeaderText = "Description"
        dgvWgs84ToOutputDTParams.Columns(6).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvWgs84ToOutputDTParams.AutoResizeColumns()


        ShowInputCrsInfo()
        ShowOutputCrsInfo()
        DisplayDirectTransformationOptions()
        DisplayInputToWgs84TransOptions()
        DisplayWgs84ToOutputTransOptions()

        'DisplayDirectTransformationOptions()
        DisplayTransformationOptions(Conversion.InputCrs, Conversion.OutputCrs)

        'Display the selected Datum Transformation methods:
        SelectDirectTransOpCode(DirectCoordOpCode)
        SelectInputToWgs84TransOpCode(InputToWgs84CoordOpCode)
        SelectWgs84ToOutputTransOpCode(Wgs84ToOutputCoordOpCode)

        If txtInputCrsQuery.Text.Trim = "" Then txtInputCrsQuery.Text = "Select COORD_REF_SYS_CODE, COORD_REF_SYS_NAME, COORD_REF_SYS_KIND, REMARKS From [Coordinate Reference System]"
        If txtOutputCrsQuery.Text.Trim = "" Then txtOutputCrsQuery.Text = "Select COORD_REF_SYS_CODE, COORD_REF_SYS_NAME, COORD_REF_SYS_KIND, REMARKS From [Coordinate Reference System]"

        ReCalcDatumTransTable()

        FindDefaultDatumTrans()

        dgvConversion.AutoResizeColumns()
        dgvInputLocations.AutoResizeColumns()
        dgvOutputLocations.AutoResizeColumns()

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form

        If FormNo > -1 Then
            Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the ChartFormClosed method to select the correct form to set to nothing.
        End If

        Me.Close() 'Close the form
    End Sub

    Private Sub Conversion_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub

    Private Sub Conversion_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If FormNo > -1 Then
            Main.ConversionFormClosed()
        End If
    End Sub

    Private Sub Conversion_ErrorMessage(Msg As String) Handles Conversion.ErrorMessage
        Main.Message.AddWarning(Msg)
    End Sub

    Private Sub Conversion_Message(Msg As String) Handles Conversion.Message
        Main.Message.Add(Msg)
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Public Sub ShowInputCrsInfo()
        'Show the Input Coordinate Reference System information.

        txtInputCrsName.Text = Conversion.InputCrs.Name
        txtInputCrsName2.Text = Conversion.InputCrs.Name
        txtInputCrsCode.Text = Conversion.InputCrs.Code
        txtInputCrsRemarks.Text = Conversion.InputCrs.Remarks

        'Show the CRS kind:
        Select Case Conversion.InputCrs.Kind
            Case CoordRefSystem.CrsKind.compound
                txtInputCrsKind.Text = "Compound"
            Case CoordRefSystem.CrsKind.derived
                txtInputCrsKind.Text = "Derived"
            Case CoordRefSystem.CrsKind.engineering
                txtInputCrsKind.Text = "Engineering"
            Case CoordRefSystem.CrsKind.geocentric
                txtInputCrsKind.Text = "Geocentric"
            Case CoordRefSystem.CrsKind.geodetic
                txtInputCrsKind.Text = "Geodetic"
            Case CoordRefSystem.CrsKind.geographic2D
                txtInputCrsKind.Text = "Geographic 2D"
            Case CoordRefSystem.CrsKind.geographic3D
                txtInputCrsKind.Text = "Geographic 3D"
            Case CoordRefSystem.CrsKind.projected
                txtInputCrsKind.Text = "Projected"
            Case CoordRefSystem.CrsKind.vertical
                txtInputCrsKind.Text = "Vertical"
            Case Else
                'txtInputCrsKind.Text = "Unrecognised: " & Conversion.InputCoordRefSystem.Kind.ToString
                txtInputCrsKind.Text = "Unrecognised: " & Conversion.InputCrs.Kind.ToString
        End Select

        'Show the CRS extent:
        txtInputCrsExtentName.Text = Conversion.InputCrs.Extent.Name
        txtInputCrsExtentDescr.Text = Conversion.InputCrs.Extent.Description
        txtInputCrsExtentRemarks.Text = Conversion.InputCrs.Extent.Remarks
        txtInputCrsExtentNorth.Text = Conversion.InputCrs.Extent.NorthBoundLat
        txtInputCrsExtentSouth.Text = Conversion.InputCrs.Extent.SouthBoundLat
        txtInputCrsExtentWest.Text = Conversion.InputCrs.Extent.WestBoundLon
        txtInputCrsExtentEast.Text = Conversion.InputCrs.Extent.EastBoundLon

        txtInputCrsCoordSysName.Text = Conversion.InputCrs.CoordSystem.Name

        udInputDatumLevel.Maximum = NBaseCrsLevels(Conversion.InputCrs, 0)
        txtInputDatumNLevels.Text = udInputDatumLevel.Maximum

        If Conversion.InputCrs.Datum.Code = -1 Then 'The InputCRS does not have a datum defined - use a BaseCrs datum if available.
            If udInputDatumLevel.Maximum = 0 Then 'There are no BaseCrs datums to use.
                GroupBox2.Text = "Datum"

                udInputDatumLevel.Enabled = False

                txtInputCrsBaseCrsName.Text = ""
                txtInputCrsBaseCrsName2.Text = ""
                txtInputCrsBaseCrsKind.Text = ""

                'Ellipsoid information:
                txtInputCrsEllipsoidName.Text = ""
                txtInputCrsSemiMajAxis.Text = ""
                txtInputCrsInvFlat.Text = ""
                txtInputCrsSemiMinAxis.Text = ""
                txtInputCrsEllipsShape.Text = ""
                txtInputCrsEllipsoidRemarks.Text = ""

                'Prime meridian information:
                txtInputCrsPMName.Text = ""
                txtInputCrsPMGreenLong.Text = ""
                txtInputCrsPMRemarks.Text = ""
            Else 'Use the first BaseCrs datum:
                GroupBox2.Text = "Base CRS Datum:"
                udInputDatumLevel.Enabled = True
                udInputDatumLevel.Minimum = 1
                udInputDatumLevel.Value = 1

                txtInputCrsBaseCrsName.Text = Conversion.InputCrs.BaseCrs.Name

                txtInputCrsBaseCrsCode.Text = Conversion.InputCrs.BaseCrs.Code

                Label55.Enabled = True
                txtInputCrsBaseCrsName2.Enabled = True
                txtInputCrsBaseCrsName2.Text = Conversion.InputCrs.BaseCrs.Name 'This BaseCrs name will be changed when the udDatumLevel selection is changed.
                Label57.Enabled = True
                txtInputCrsBaseCrsCode2.Enabled = True
                txtInputCrsBaseCrsCode2.Text = Conversion.InputCrs.BaseCrs.Code
                'txtInputCrsBaseCrsKind.Text = Conversion.InputCRS.BaseCrs.Kind
                Label56.Enabled = True
                txtInputCrsBaseCrsKind.Enabled = True
                Select Case Conversion.InputCrs.BaseCrs.Kind
                    Case CoordRefSystem.CrsKind.compound
                        txtInputCrsBaseCrsKind.Text = "Compound"
                    Case CoordRefSystem.CrsKind.derived
                        txtInputCrsBaseCrsKind.Text = "Derived"
                    Case CoordRefSystem.CrsKind.engineering
                        txtInputCrsBaseCrsKind.Text = "Engineering"
                    Case CoordRefSystem.CrsKind.geocentric
                        txtInputCrsBaseCrsKind.Text = "Geocentric"
                    Case CoordRefSystem.CrsKind.geodetic
                        txtInputCrsBaseCrsKind.Text = "Geodetic"
                    Case CoordRefSystem.CrsKind.geographic2D
                        txtInputCrsBaseCrsKind.Text = "Geographic 2D"
                    Case CoordRefSystem.CrsKind.geographic3D
                        txtInputCrsBaseCrsKind.Text = "Geographic 3D"
                    Case CoordRefSystem.CrsKind.projected
                        txtInputCrsBaseCrsKind.Text = "Projected"
                    Case CoordRefSystem.CrsKind.vertical
                        txtInputCrsBaseCrsKind.Text = "Vertical"
                    Case Else
                        txtInputCrsBaseCrsKind.Text = "Unrecognised: " & Conversion.InputCrs.Kind.ToString
                End Select

                'Datum name:
                txtInputCrsDatumName.Text = Conversion.InputCrs.BaseCrs.Datum.Name

                'Ellipsoid information:
                txtInputCrsEllipsoidName.Text = Conversion.InputCrs.BaseCrs.Ellipsoid.Name
                txtInputCrsSemiMajAxis.Text = Conversion.InputCrs.BaseCrs.Ellipsoid.SemiMajorAxis
                txtInputCrsInvFlat.Text = Conversion.InputCrs.BaseCrs.Ellipsoid.InvFlattening
                txtInputCrsSemiMinAxis.Text = Conversion.InputCrs.BaseCrs.Ellipsoid.SemiMinorAxis
                txtInputCrsEllipsShape.Text = Conversion.InputCrs.BaseCrs.Ellipsoid.EllipsoidShape
                txtInputCrsEllipsoidRemarks.Text = Conversion.InputCrs.BaseCrs.Ellipsoid.Remarks

                'Prime meridian information:
                txtInputCrsPMName.Text = Conversion.InputCrs.BaseCrs.PrimeMeridian.Name
                txtInputCrsPMGreenLong.Text = Conversion.InputCrs.BaseCrs.PrimeMeridian.GreenwichLongitude
                txtInputCrsPMRemarks.Text = Conversion.InputCrs.BaseCrs.PrimeMeridian.Remarks
            End If
        Else
            GroupBox2.Text = "Datum"
            Label55.Enabled = False
            txtInputCrsBaseCrsName2.Enabled = False
            txtInputCrsBaseCrsName2.Text = "" 'A BaseCrs Datum is not being displayed.
            Label57.Enabled = False
            txtInputCrsBaseCrsCode2.Enabled = False
            txtInputCrsBaseCrsCode2.Text = ""
            Label56.Enabled = False
            txtInputCrsBaseCrsKind.Enabled = False
            txtInputCrsBaseCrsKind.Text = ""

            udInputDatumLevel.Enabled = True
            udInputDatumLevel.Minimum = 0
            udInputDatumLevel.Value = 0

            If IsNothing(Conversion.InputCrs.BaseCrs) Then
                txtInputCrsBaseCrsName.Text = ""
                txtInputCrsBaseCrsCode.Text = ""
            Else
                txtInputCrsBaseCrsName.Text = Conversion.InputCrs.BaseCrs.Name
                txtInputCrsBaseCrsCode.Text = Conversion.InputCrs.BaseCrs.Code
            End If



            'Datum name:
            txtInputCrsDatumName.Text = Conversion.InputCrs.Datum.Name

            'Ellipsoid information:
            txtInputCrsEllipsoidName.Text = Conversion.InputCrs.Ellipsoid.Name
            txtInputCrsSemiMajAxis.Text = Conversion.InputCrs.Ellipsoid.SemiMajorAxis
            txtInputCrsInvFlat.Text = Conversion.InputCrs.Ellipsoid.InvFlattening
            txtInputCrsSemiMinAxis.Text = Conversion.InputCrs.Ellipsoid.SemiMinorAxis
            txtInputCrsEllipsShape.Text = Conversion.InputCrs.Ellipsoid.EllipsoidShape
            txtInputCrsEllipsoidRemarks.Text = Conversion.InputCrs.Ellipsoid.Remarks

            'Prime meridian information:
            txtInputCrsPMName.Text = Conversion.InputCrs.PrimeMeridian.Name
            txtInputCrsPMGreenLong.Text = Conversion.InputCrs.PrimeMeridian.GreenwichLongitude
            txtInputCrsPMRemarks.Text = Conversion.InputCrs.PrimeMeridian.Remarks
        End If

        dgvInputSourceCoordOps.Rows.Clear()
        dgvInputTargetCoordOps.Rows.Clear()
        'ListInputSourceCoordOps(Conversion.InputCrs, 0)
        ListInputSourceTargetCoordOps(Conversion.InputCrs, 0)
        dgvInputSourceCoordOps.AutoResizeColumns()
        dgvInputTargetCoordOps.AutoResizeColumns()

        Dim CoordOpSteps As Integer = Conversion.InputCrs.DefiningCoordOpList.Count

        'Coordinate System:
        txtInputCrsCoordSysName.Text = Conversion.InputCrs.CoordSystem.Name
        txtInputCrsCoordSysCode.Text = Conversion.InputCrs.CoordSystem.Code
        txtInputCrsCoordSysType.Text = Conversion.InputCrs.CoordSystem.Type.ToString
        txtInputCrsCoordSysDim.Text = Conversion.InputCrs.CoordSystem.Dimension

        If Conversion.InputCrs.CoordSystem.Dimension = 0 Then
            txtInputCrsCoordSysAx1Name.Enabled = False
            txtInputCrsCoordSysAx1Name.Text = ""
            rbInputCrsCoordSysAx1Name.Enabled = False
            rbInputCrsCoordSysAx1Name.Checked = False
            txtInputCrsCoordSysAx2Name.Enabled = False
            txtInputCrsCoordSysAx2Name.Text = ""
            rbInputCrsCoordSysAx2Name.Enabled = False
            rbInputCrsCoordSysAx2Name.Checked = False
            txtInputCrsCoordSysAx3Name.Enabled = False
            txtInputCrsCoordSysAx3Name.Text = ""
            rbInputCrsCoordSysAx3Name.Enabled = False
            rbInputCrsCoordSysAx3Name.Checked = False
            DisplayInputAxisInfo(0)
        ElseIf Conversion.InputCrs.CoordSystem.Dimension = 1 Then
            txtInputCrsCoordSysAx1Name.Enabled = True
            If Conversion.InputCrs.CoordAxisNameList.Count = 0 Then
                Main.Message.AddWarning("The Input CRS coord system has 1 dimension but the axis name list is empty." & vbCrLf)
            Else
                txtInputCrsCoordSysAx1Name.Text = Conversion.InputCrs.CoordAxisNameList(0).Name
                rbInputCrsCoordSysAx1Name.Enabled = True
                rbInputCrsCoordSysAx1Name.Checked = True
            End If
            txtInputCrsCoordSysAx2Name.Text = ""
            rbInputCrsCoordSysAx2Name.Enabled = False
            rbInputCrsCoordSysAx2Name.Checked = False
            txtInputCrsCoordSysAx3Name.Enabled = False
            txtInputCrsCoordSysAx3Name.Text = ""
            rbInputCrsCoordSysAx3Name.Enabled = False
            rbInputCrsCoordSysAx3Name.Checked = False
            DisplayInputAxisInfo(1) 'Display information about the first axis.
        ElseIf Conversion.InputCrs.CoordSystem.Dimension = 2 Then
            If Conversion.InputCrs.CoordAxisNameList.Count = 0 Then
                Main.Message.AddWarning("The Input CRS coord system has 2 dimensions but the axis name list is empty." & vbCrLf)
            Else
                txtInputCrsCoordSysAx1Name.Enabled = True
                txtInputCrsCoordSysAx1Name.Text = Conversion.InputCrs.CoordAxisNameList(0).Name
                rbInputCrsCoordSysAx1Name.Enabled = True
                rbInputCrsCoordSysAx1Name.Checked = True
                txtInputCrsCoordSysAx2Name.Text = Conversion.InputCrs.CoordAxisNameList(1).Name
                rbInputCrsCoordSysAx2Name.Enabled = True
                txtInputCrsCoordSysAx3Name.Enabled = False
            End If
            txtInputCrsCoordSysAx3Name.Text = ""
            rbInputCrsCoordSysAx3Name.Enabled = False
            DisplayInputAxisInfo(1) 'Display information about the first axis.
        ElseIf Conversion.InputCrs.CoordSystem.Dimension = 3 Then
            If Conversion.InputCrs.CoordAxisNameList.Count = 0 Then
                Main.Message.AddWarning("The Input CRS coord system has 3 dimensions but the axis name list is empty." & vbCrLf)
            Else
                txtInputCrsCoordSysAx1Name.Enabled = True
                txtInputCrsCoordSysAx1Name.Text = Conversion.InputCrs.CoordAxisNameList(0).Name
                rbInputCrsCoordSysAx1Name.Enabled = True
                rbInputCrsCoordSysAx1Name.Checked = True
                txtInputCrsCoordSysAx2Name.Text = Conversion.InputCrs.CoordAxisNameList(1).Name
                rbInputCrsCoordSysAx2Name.Enabled = True
                txtInputCrsCoordSysAx3Name.Enabled = True
                txtInputCrsCoordSysAx3Name.Text = Conversion.InputCrs.CoordAxisNameList(2).Name
                rbInputCrsCoordSysAx3Name.Enabled = True
            End If
            DisplayInputAxisInfo(1) 'Display information about the first axis.
        Else

        End If

        'Projection operation used to convert between the Derived CRS and the Base CRS.:
        If Conversion.InputCrs.ProjConvCode = -1 Then
            'No Projection Conversion required.
            txtInputCrsProjConvName.Text = ""
            txtInputCrsProjMethodName.Text = ""
            txtInputCrsProjMethodName2.Text = ""
            txtInputCrsProjMethodReversable.Text = ""
            txtInputCrsProjMethodFormula.Text = ""
            txtInputCrsProjMethodExample.Text = ""
            txtInputCrsProjMethodCode.Text = ""
            txtInputCrsProjMethodRemarks.Text = ""
            dgvInputCrsProjParams.Rows.Clear()
        Else
            txtInputCrsProjConvName.Text = Conversion.InputCrs.ProjectionCoordOp.Name
            txtInputCrsProjMethodName.Text = Conversion.InputCrs.ProjectionCoordOpMethod.Name
            txtInputCrsProjMethodName2.Text = Conversion.InputCrs.ProjectionCoordOpMethod.Name
            txtInputCrsProjMethodReversable.Text = Conversion.InputCrs.ProjectionCoordOpMethod.ReverseOp.ToString
            txtInputCrsProjMethodFormula.Text = Conversion.InputCrs.ProjectionCoordOpMethod.Formula
            txtInputCrsProjMethodExample.Text = Conversion.InputCrs.ProjectionCoordOpMethod.Example
            txtInputCrsProjMethodCode.Text = Conversion.InputCrs.ProjectionCoordOpMethod.Code
            txtInputCrsProjMethodRemarks.Text = Conversion.InputCrs.ProjectionCoordOpMethod.Remarks
            dgvInputCrsProjParams.Rows.Clear()

            Dim ParamNo As Integer
            Dim NParams As Integer = Conversion.InputCrs.ProjectionCoordOpParamList.Count
            For ParamNo = 0 To NParams - 1
                dgvInputCrsProjParams.Rows.Add(Conversion.InputCrs.ProjectionCoordOpParamList(ParamNo).Code,
                                               Conversion.InputCrs.ProjectionCoordOpParamUseList(ParamNo).SortOrder,
                                               Conversion.InputCrs.ProjectionCoordOpParamList(ParamNo).Name,
                                               Conversion.InputCrs.ProjectionCoordOpParamUseList(ParamNo).SignReversal,
                                               Conversion.InputCrs.ProjectionCoordOpParamValList(ParamNo).ParameterValue,
                                               Conversion.UnitOfMeas(Conversion.InputCrs.ProjectionCoordOpParamValList(ParamNo).UomCode).Name,
                                               Conversion.InputCrs.ProjectionCoordOpParamList(ParamNo).Description)
            Next
            dgvInputCrsProjParams.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            dgvInputCrsProjParams.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvInputCrsProjParams.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvInputCrsProjParams.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvInputCrsProjParams.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvInputCrsProjParams.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvInputCrsProjParams.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

            dgvInputCrsProjParams.Columns(6).Width = 640

            dgvInputCrsProjParams.Columns(6).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvInputCrsProjParams.AutoResizeRows()

            dgvInputCrsProjParams.AllowUserToAddRows = False
        End If
    End Sub

    'Private Sub UpdateInputCrsTypeConversion()
    '    'Update the settings in Input CRS Type Conversion - On the Point Conversion \ Coordinate Type Conversion tab.

    '    If Conversion.InputCrs.Kind = "Projected" Then
    '        rbInputEastNorth.Enabled = True
    '    Else
    '        rbInputEastNorth.Enabled = False
    '        If rbInputEastNorth.Checked Then
    '            rbInputLongLat.Enabled = True
    '        End If
    '    End If

    'End Sub

    Public Sub ShowOutputCrsInfo()
        'Show the Output Coordinate Reference System information.

        txtOutputCrsName.Text = Conversion.OutputCrs.Name
        txtOutputCrsName2.Text = Conversion.OutputCrs.Name
        txtOutputCrsCode.Text = Conversion.OutputCrs.Code
        txtOutputCrsRemarks.Text = Conversion.OutputCrs.Remarks

        'Show the CRS kind:
        Select Case Conversion.OutputCrs.Kind
            Case CoordRefSystem.CrsKind.compound
                txtOutputCrsKind.Text = "Compound"
            Case CoordRefSystem.CrsKind.derived
                txtOutputCrsKind.Text = "Derived"
            Case CoordRefSystem.CrsKind.engineering
                txtOutputCrsKind.Text = "Engineering"
            Case CoordRefSystem.CrsKind.geocentric
                txtOutputCrsKind.Text = "Geocentric"
            Case CoordRefSystem.CrsKind.geodetic
                txtOutputCrsKind.Text = "Geodetic"
            Case CoordRefSystem.CrsKind.geographic2D
                txtOutputCrsKind.Text = "Geographic 2D"
            Case CoordRefSystem.CrsKind.geographic3D
                txtOutputCrsKind.Text = "Geographic 3D"
            Case CoordRefSystem.CrsKind.projected
                txtOutputCrsKind.Text = "Projected"
            Case CoordRefSystem.CrsKind.vertical
                txtOutputCrsKind.Text = "Vertical"
            Case Else
                txtOutputCrsKind.Text = "Unrecognised: " & Conversion.OutputCrs.Kind.ToString
        End Select

        'Show the CRS extent:
        txtOutputCrsExtentName.Text = Conversion.OutputCrs.Extent.Name
        txtOutputCrsExtentDescr.Text = Conversion.OutputCrs.Extent.Description
        txtOutputCrsExtentRemarks.Text = Conversion.OutputCrs.Extent.Remarks
        txtOutputCrsExtentNorth.Text = Conversion.OutputCrs.Extent.NorthBoundLat
        txtOutputCrsExtentSouth.Text = Conversion.OutputCrs.Extent.SouthBoundLat
        txtOutputCrsExtentWest.Text = Conversion.OutputCrs.Extent.WestBoundLon
        txtOutputCrsExtentEast.Text = Conversion.OutputCrs.Extent.EastBoundLon

        txtOutputCrsCoordSysName.Text = Conversion.OutputCrs.CoordSystem.Name

        udOutputDatumLevel.Maximum = NBaseCrsLevels(Conversion.OutputCrs, 0)
        txtOutputDatumNLevels.Text = udOutputDatumLevel.Maximum

        If Conversion.OutputCrs.Datum.Code = -1 Then 'The OutputCRS does not have a datum defined - use a BaseCrs datum if available.
            If udOutputDatumLevel.Maximum = 0 Then 'There are no BaseCrs datume to use.
                GroupBox10.Text = "Datum"
                udOutputDatumLevel.Enabled = False

                txtOutputCrsBaseCrsName.Text = ""
                txtOutputCrsBaseCrsName2.Text = ""
                txtOutputCrsBaseCrsKind.Text = ""

                'Ellipsoid information:
                txtOutputCrsEllipsoidName.Text = ""
                txtOutputCrsSemiMajAxis.Text = ""
                txtOutputCrsInvFlat.Text = ""
                txtOutputCrsSemiMinAxis.Text = ""
                txtOutputCrsEllipsShape.Text = ""
                txtOutputCrsEllipsoidRemarks.Text = ""

                'Prime meridian information:
                txtOutputCrsPMName.Text = ""
                txtOutputCrsPMGreenLong.Text = ""
                txtOutputCrsPMRemarks.Text = ""
            Else 'Use the first BaseCrs datum:
                GroupBox10.Text = "Base CRS Datum:"
                udOutputDatumLevel.Enabled = True
                udOutputDatumLevel.Minimum = 1
                udOutputDatumLevel.Value = 1

                txtOutputCrsBaseCrsName.Text = Conversion.OutputCrs.BaseCrs.Name
                txtOutputCrsBaseCrsCode.Text = Conversion.OutputCrs.BaseCrs.Code

                Label76.Enabled = True
                txtOutputCrsBaseCrsName2.Enabled = True
                txtOutputCrsBaseCrsName2.Text = Conversion.OutputCrs.BaseCrs.Name 'This BaseCrs name will be changed when the udDatumLevel selection is changed.
                Label74.Enabled = True
                txtOutputCrsBaseCrsCode2.Enabled = True
                txtOutputCrsBaseCrsCode2.Text = Conversion.OutputCrs.BaseCrs.Code
                'txtInputCrsBaseCrsKind.Text = Conversion.InputCRS.BaseCrs.Kind
                Label75.Enabled = True
                txtOutputCrsBaseCrsKind.Enabled = True
                Select Case Conversion.OutputCrs.BaseCrs.Kind
                    Case CoordRefSystem.CrsKind.compound
                        txtOutputCrsBaseCrsKind.Text = "Compound"
                    Case CoordRefSystem.CrsKind.derived
                        txtOutputCrsBaseCrsKind.Text = "Derived"
                    Case CoordRefSystem.CrsKind.engineering
                        txtOutputCrsBaseCrsKind.Text = "Engineering"
                    Case CoordRefSystem.CrsKind.geocentric
                        txtOutputCrsBaseCrsKind.Text = "Geocentric"
                    Case CoordRefSystem.CrsKind.geodetic
                        txtOutputCrsBaseCrsKind.Text = "Geodetic"
                    Case CoordRefSystem.CrsKind.geographic2D
                        txtOutputCrsBaseCrsKind.Text = "Geographic 2D"
                    Case CoordRefSystem.CrsKind.geographic3D
                        txtOutputCrsBaseCrsKind.Text = "Geographic 3D"
                    Case CoordRefSystem.CrsKind.projected
                        txtOutputCrsBaseCrsKind.Text = "Projected"
                    Case CoordRefSystem.CrsKind.vertical
                        txtOutputCrsBaseCrsKind.Text = "Vertical"
                    Case Else
                        txtOutputCrsBaseCrsKind.Text = "Unrecognised: " & Conversion.OutputCrs.Kind.ToString
                End Select

                'Datum name:
                txtOutputCrsDatumName.Text = Conversion.OutputCrs.BaseCrs.Datum.Name

                'Ellipsoid information:
                txtOutputCrsEllipsoidName.Text = Conversion.OutputCrs.BaseCrs.Ellipsoid.Name
                txtOutputCrsSemiMajAxis.Text = Conversion.OutputCrs.BaseCrs.Ellipsoid.SemiMajorAxis
                txtOutputCrsInvFlat.Text = Conversion.OutputCrs.BaseCrs.Ellipsoid.InvFlattening
                txtOutputCrsSemiMinAxis.Text = Conversion.OutputCrs.BaseCrs.Ellipsoid.SemiMinorAxis
                txtOutputCrsEllipsShape.Text = Conversion.OutputCrs.BaseCrs.Ellipsoid.EllipsoidShape
                txtOutputCrsEllipsoidRemarks.Text = Conversion.OutputCrs.BaseCrs.Ellipsoid.Remarks

                'Prime meridian information:
                txtOutputCrsPMName.Text = Conversion.OutputCrs.BaseCrs.PrimeMeridian.Name
                txtOutputCrsPMGreenLong.Text = Conversion.OutputCrs.BaseCrs.PrimeMeridian.GreenwichLongitude
                txtOutputCrsPMRemarks.Text = Conversion.OutputCrs.BaseCrs.PrimeMeridian.Remarks
            End If
        Else 'Show the OutputCRS datum
            GroupBox2.Text = "Datum"
            Label76.Enabled = False
            txtOutputCrsBaseCrsName2.Enabled = False
            txtOutputCrsBaseCrsName2.Text = "" 'A BaseCrs Datum is not being displayed.
            Label74.Enabled = False
            txtOutputCrsBaseCrsCode2.Enabled = False
            txtOutputCrsBaseCrsCode2.Text = ""
            Label75.Enabled = False
            txtOutputCrsBaseCrsKind.Enabled = False
            txtOutputCrsBaseCrsKind.Text = ""


            udOutputDatumLevel.Enabled = True
            udOutputDatumLevel.Minimum = 0
            udOutputDatumLevel.Value = 0

            If IsNothing(Conversion.OutputCrs.BaseCrs) Then
                txtOutputCrsBaseCrsName.Text = ""
                txtOutputCrsBaseCrsCode.Text = ""
            Else
                txtOutputCrsBaseCrsName.Text = Conversion.OutputCrs.BaseCrs.Name
                txtOutputCrsBaseCrsCode.Text = Conversion.OutputCrs.BaseCrs.Code
            End If



            'Datum name:
            txtOutputCrsDatumName.Text = Conversion.OutputCrs.Datum.Name

            'Ellipsoid information:
            txtOutputCrsEllipsoidName.Text = Conversion.OutputCrs.Ellipsoid.Name
            txtOutputCrsSemiMajAxis.Text = Conversion.OutputCrs.Ellipsoid.SemiMajorAxis
            txtOutputCrsInvFlat.Text = Conversion.OutputCrs.Ellipsoid.InvFlattening
            txtOutputCrsSemiMinAxis.Text = Conversion.OutputCrs.Ellipsoid.SemiMinorAxis
            txtOutputCrsEllipsShape.Text = Conversion.OutputCrs.Ellipsoid.EllipsoidShape
            txtOutputCrsEllipsoidRemarks.Text = Conversion.OutputCrs.Ellipsoid.Remarks

            'Prime meridian information:
            txtOutputCrsPMName.Text = Conversion.OutputCrs.PrimeMeridian.Name
            txtOutputCrsPMGreenLong.Text = Conversion.OutputCrs.PrimeMeridian.GreenwichLongitude
            txtOutputCrsPMRemarks.Text = Conversion.OutputCrs.PrimeMeridian.Remarks
        End If

        dgvOutputSourceCoordOps.Rows.Clear()
        dgvOutputTargetCoordOps.Rows.Clear()
        'ListOutputTargetCoordOps(Conversion.OutputCrs, 0)
        ListOutputSourceTargetCoordOps(Conversion.OutputCrs, 0)
        dgvOutputSourceCoordOps.AutoResizeColumns()
        dgvOutputTargetCoordOps.AutoResizeColumns()

        Dim CoordOpSteps As Integer = Conversion.OutputCrs.DefiningCoordOpList.Count

        'Coordinate System:
        txtOutputCrsCoordSysName.Text = Conversion.OutputCrs.CoordSystem.Name
        txtOutputCrsCoordSysCode.Text = Conversion.OutputCrs.CoordSystem.Code
        txtOutputCrsCoordSysType.Text = Conversion.OutputCrs.CoordSystem.Type.ToString
        txtOutputCrsCoordSysDim.Text = Conversion.OutputCrs.CoordSystem.Dimension

        If Conversion.OutputCrs.CoordSystem.Dimension = 0 Then
            txtOutputCrsCoordSysAx1Name.Enabled = False
            txtOutputCrsCoordSysAx1Name.Text = ""
            rbOutputCrsCoordSysAx1Name.Enabled = False
            rbOutputCrsCoordSysAx1Name.Checked = False
            txtOutputCrsCoordSysAx2Name.Enabled = False
            txtOutputCrsCoordSysAx2Name.Text = ""
            rbOutputCrsCoordSysAx2Name.Enabled = False
            rbOutputCrsCoordSysAx2Name.Checked = False
            txtOutputCrsCoordSysAx3Name.Enabled = False
            txtOutputCrsCoordSysAx3Name.Text = ""
            rbOutputCrsCoordSysAx3Name.Enabled = False
            rbOutputCrsCoordSysAx3Name.Checked = False
            DisplayOutputAxisInfo(0)
        ElseIf Conversion.OutputCrs.CoordSystem.Dimension = 1 Then
            If Conversion.OutputCrs.CoordAxisNameList.Count = 0 Then
                Main.Message.AddWarning("The Output CRS coord system has 1 dimension but the axis name list is empty." & vbCrLf)
            Else
                txtOutputCrsCoordSysAx1Name.Enabled = True
                txtOutputCrsCoordSysAx1Name.Text = Conversion.OutputCrs.CoordAxisNameList(0).Name
                rbOutputCrsCoordSysAx1Name.Enabled = True
                rbOutputCrsCoordSysAx1Name.Checked = True
            End If
            txtOutputCrsCoordSysAx2Name.Text = ""
            rbOutputCrsCoordSysAx2Name.Enabled = False
            rbOutputCrsCoordSysAx2Name.Checked = False
            txtOutputCrsCoordSysAx3Name.Enabled = False
            txtOutputCrsCoordSysAx3Name.Text = ""
            rbOutputCrsCoordSysAx3Name.Enabled = False
            rbOutputCrsCoordSysAx3Name.Checked = False
            DisplayOutputAxisInfo(1) 'Display information about the first axis.
        ElseIf Conversion.OutputCrs.CoordSystem.Dimension = 2 Then
            If Conversion.OutputCrs.CoordAxisNameList.Count = 0 Then
                Main.Message.AddWarning("The Output CRS coord system has 2 dimensions but the axis name list is empty." & vbCrLf)
            Else
                txtOutputCrsCoordSysAx1Name.Enabled = True
                txtOutputCrsCoordSysAx1Name.Text = Conversion.OutputCrs.CoordAxisNameList(0).Name
                rbOutputCrsCoordSysAx1Name.Enabled = True
                rbOutputCrsCoordSysAx1Name.Checked = True
                txtOutputCrsCoordSysAx2Name.Text = Conversion.OutputCrs.CoordAxisNameList(1).Name
                rbOutputCrsCoordSysAx2Name.Enabled = True
            End If
            txtOutputCrsCoordSysAx3Name.Enabled = False
            txtOutputCrsCoordSysAx3Name.Text = ""
            rbOutputCrsCoordSysAx3Name.Enabled = False
            DisplayOutputAxisInfo(1) 'Display information about the first axis.
        ElseIf Conversion.OutputCrs.CoordSystem.Dimension = 3 Then
            If Conversion.OutputCrs.CoordAxisNameList.Count = 0 Then
                Main.Message.AddWarning("The Output CRS coord system has 3 dimensions but the axis name list is empty." & vbCrLf)
            Else
                txtOutputCrsCoordSysAx1Name.Enabled = True
                txtOutputCrsCoordSysAx1Name.Text = Conversion.OutputCrs.CoordAxisNameList(0).Name
                rbOutputCrsCoordSysAx1Name.Enabled = True
                rbOutputCrsCoordSysAx1Name.Checked = True
                txtOutputCrsCoordSysAx2Name.Text = Conversion.OutputCrs.CoordAxisNameList(1).Name
                rbOutputCrsCoordSysAx2Name.Enabled = True
                txtOutputCrsCoordSysAx3Name.Enabled = True
                txtOutputCrsCoordSysAx3Name.Text = Conversion.OutputCrs.CoordAxisNameList(2).Name
                rbOutputCrsCoordSysAx3Name.Enabled = True
            End If
            DisplayOutputAxisInfo(1) 'Display information about the first axis.
        Else

        End If

        'Projection operation used to convert between the Derived CRS and the Base CRS.:
        If Conversion.OutputCrs.ProjConvCode = -1 Then
            'No Projection Conversion required.
            txtOutputCrsProjConvName.Text = ""
            txtOutputCrsProjMethodName.Text = ""
            txtOutputCrsProjMethodName2.Text = ""
            txtOutputCrsProjMethodReversable.Text = ""
            txtOutputCrsProjMethodFormula.Text = ""
            txtOutputCrsProjMethodExample.Text = ""
            txtOutputCrsProjMethodCode.Text = ""
            txtOutputCrsProjMethodRemarks.Text = ""
            dgvOutputCrsProjParams.Rows.Clear()
        Else
            txtOutputCrsProjConvName.Text = Conversion.OutputCrs.ProjectionCoordOp.Name
            txtOutputCrsProjMethodName.Text = Conversion.OutputCrs.ProjectionCoordOpMethod.Name
            txtOutputCrsProjMethodName2.Text = Conversion.OutputCrs.ProjectionCoordOpMethod.Name
            txtOutputCrsProjMethodReversable.Text = Conversion.OutputCrs.ProjectionCoordOpMethod.ReverseOp.ToString
            txtOutputCrsProjMethodFormula.Text = Conversion.OutputCrs.ProjectionCoordOpMethod.Formula
            txtOutputCrsProjMethodExample.Text = Conversion.OutputCrs.ProjectionCoordOpMethod.Example
            txtOutputCrsProjMethodCode.Text = Conversion.OutputCrs.ProjectionCoordOpMethod.Code
            txtOutputCrsProjMethodRemarks.Text = Conversion.OutputCrs.ProjectionCoordOpMethod.Remarks
            dgvOutputCrsProjParams.Rows.Clear()

            Dim ParamNo As Integer
            Dim NParams As Integer = Conversion.OutputCrs.ProjectionCoordOpParamList.Count
            For ParamNo = 0 To NParams - 1
                dgvOutputCrsProjParams.Rows.Add(Conversion.OutputCrs.ProjectionCoordOpParamList(ParamNo).Code,
                                               Conversion.OutputCrs.ProjectionCoordOpParamUseList(ParamNo).SortOrder,
                                               Conversion.OutputCrs.ProjectionCoordOpParamList(ParamNo).Name,
                                               Conversion.OutputCrs.ProjectionCoordOpParamUseList(ParamNo).SignReversal,
                                               Conversion.OutputCrs.ProjectionCoordOpParamValList(ParamNo).ParameterValue,
                                               Conversion.UnitOfMeas(Conversion.OutputCrs.ProjectionCoordOpParamValList(ParamNo).UomCode).Name,
                                               Conversion.OutputCrs.ProjectionCoordOpParamList(ParamNo).Description)
            Next
            dgvOutputCrsProjParams.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            dgvOutputCrsProjParams.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvOutputCrsProjParams.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvOutputCrsProjParams.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvOutputCrsProjParams.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvOutputCrsProjParams.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvOutputCrsProjParams.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

            dgvOutputCrsProjParams.Columns(6).Width = 640

            dgvOutputCrsProjParams.Columns(6).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvOutputCrsProjParams.AutoResizeRows()

            dgvOutputCrsProjParams.AllowUserToAddRows = False
        End If
    End Sub

    'Private Sub ListInputSourceCoordOps(CRS As CoordRefSystem, Level As Integer)
    Private Sub ListInputSourceTargetCoordOps(CRS As CoordRefSystem, Level As Integer)
        'Display the list of Input Coordinate Operations that use the Input CRS as the Source or the Target.
        'Level is the list Level: 0 for the CRS list, 1 for the Base CRS list 2 for the Base of the Base CRS list etc.

        For Each Item As CoordinateOperation In CRS.SourceCoordOpList
            dgvInputSourceCoordOps.Rows.Add(Level, Item.Name, Item.Type, Item.Code, Item.Accuracy.ToString, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsCode, Item.TargetCrsCode)
        Next

        For Each Item As CoordinateOperation In CRS.TargetCoordOpList
            dgvInputTargetCoordOps.Rows.Add(Level, Item.Name, Item.Type, Item.Code, Item.Accuracy.ToString, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsCode, Item.TargetCrsCode)
        Next

        If IsNothing(CRS.BaseCrs) Then
            'There is no Base CRS to process.
            Main.Message.Add("CRS Code " & CRS.Code & " does not have a base CRS." & vbCrLf)
        Else
            'ListInputSourceCoordOps(CRS.BaseCrs, Level + 1) 'Display the list of source coord operation in the Base CRS
            ListInputSourceTargetCoordOps(CRS.BaseCrs, Level + 1) 'Display the list of source coord operation in the Base CRS
            'Main.Message.Add("CRS Code " & CRS.Code & " has the Base CRS Code " & CRS.BaseCrs.Code & vbCrLf)
        End If
    End Sub

    'Private Sub ListOutputTargetCoordOps(CRS As CoordRefSystem, Level As Integer)
    Private Sub ListOutputSourceTargetCoordOps(CRS As CoordRefSystem, Level As Integer)
        'Display the list of Output Coordinate Operations the use the Output CRS as the Source or Target.
        'Level is the list Level: 0 for the CRS list, 1 for the Base CRS list 2 for the Base of the Base CRS list etc.

        For Each Item As CoordinateOperation In CRS.SourceCoordOpList
            dgvOutputSourceCoordOps.Rows.Add(Level, Item.Name, Item.Type, Item.Code, Item.Accuracy.ToString, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsCode, Item.TargetCrsCode)
        Next

        For Each Item As CoordinateOperation In CRS.TargetCoordOpList
            dgvOutputTargetCoordOps.Rows.Add(Level, Item.Name, Item.Type, Item.Code, Item.Accuracy.ToString, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsCode, Item.TargetCrsCode)
        Next

        If IsNothing(CRS.BaseCrs) Then
            'There is no Base CRS to process.
            Main.Message.Add("CRS Code " & CRS.Code & " does not have a base CRS." & vbCrLf)
        Else
            'ListOutputTargetCoordOps(CRS.BaseCrs, Level + 1) 'Display the list of source coord operation in the Base CRS
            ListOutputSourceTargetCoordOps(CRS.BaseCrs, Level + 1) 'Display the list of source coord operation in the Base CRS
        End If
    End Sub

    Private Function NBaseCrsLevels(CRS As CoordRefSystem, Level As Integer) As Integer
        'Return the number of BaseCrs levels
        If IsNothing(CRS.BaseCrs) Then
            Return Level
        Else
            Return NBaseCrsLevels(CRS.BaseCrs, Level + 1)
        End If
    End Function

    Private Sub udInputDatumLevel_ValueChanged(sender As Object, e As EventArgs) Handles udInputDatumLevel.ValueChanged
        'Display the Datum information for the selected Input BaseCrs level.
        If udInputDatumLevel.Focused Then
            DisplayInputBaseCrsLevelDatum(Conversion.InputCrs, udInputDatumLevel.Value, 0)
        End If
    End Sub

    Private Sub udOutputDatumLevel_ValueChanged(sender As Object, e As EventArgs) Handles udOutputDatumLevel.ValueChanged
        'Display the Datum information for the selected Output BaseCrs level.
        If udOutputDatumLevel.Focused Then
            DisplayOutputBaseCrsLevelDatum(Conversion.OutputCrs, udOutputDatumLevel.Value, 0)
        End If
    End Sub

    Private Sub DisplayInputBaseCrsLevelDatum(CRS As CoordRefSystem, Level As Integer, Count As Integer)
        'Display the selected BaseCrs level
        If Count >= Level Then
            'Datum name:
            txtInputCrsDatumName.Text = CRS.Datum.Name

            If Level = 0 Then
                GroupBox2.Text = "Datum:"
                Label55.Enabled = False
                txtInputCrsBaseCrsName2.Enabled = False
                txtInputCrsBaseCrsName2.Text = "" 'A BaseCrs Datum is not being displayed.
                Label57.Enabled = False
                txtInputCrsBaseCrsCode2.Enabled = False
                txtInputCrsBaseCrsCode2.Text = ""
                Label56.Enabled = False
                txtInputCrsBaseCrsKind.Enabled = False
                txtInputCrsBaseCrsKind.Text = ""
            Else
                GroupBox2.Text = "Base CRS Datum:"
                Label55.Enabled = True
                txtInputCrsBaseCrsName2.Enabled = True
                txtInputCrsBaseCrsName2.Text = CRS.Name
                Label57.Enabled = True
                txtInputCrsBaseCrsCode2.Enabled = True
                txtInputCrsBaseCrsCode2.Text = CRS.Code
                Label56.Enabled = True
                txtInputCrsBaseCrsKind.Enabled = True
                Select Case CRS.Kind
                    Case CoordRefSystem.CrsKind.compound
                        txtInputCrsBaseCrsKind.Text = "Compound"
                    Case CoordRefSystem.CrsKind.derived
                        txtInputCrsBaseCrsKind.Text = "Derived"
                    Case CoordRefSystem.CrsKind.engineering
                        txtInputCrsBaseCrsKind.Text = "Engineering"
                    Case CoordRefSystem.CrsKind.geocentric
                        txtInputCrsBaseCrsKind.Text = "Geocentric"
                    Case CoordRefSystem.CrsKind.geodetic
                        txtInputCrsBaseCrsKind.Text = "Geodetic"
                    Case CoordRefSystem.CrsKind.geographic2D
                        txtInputCrsBaseCrsKind.Text = "Geographic 2D"
                    Case CoordRefSystem.CrsKind.geographic3D
                        txtInputCrsBaseCrsKind.Text = "Geographic 3D"
                    Case CoordRefSystem.CrsKind.projected
                        txtInputCrsBaseCrsKind.Text = "Projected"
                    Case CoordRefSystem.CrsKind.vertical
                        txtInputCrsBaseCrsKind.Text = "Vertical"
                    Case Else
                        txtInputCrsBaseCrsKind.Text = "Unrecognised: " & Conversion.InputCrs.Kind.ToString
                End Select
            End If

            'Ellipsoid information:
            txtInputCrsEllipsoidName.Text = CRS.Ellipsoid.Name
            txtInputCrsSemiMajAxis.Text = CRS.Ellipsoid.SemiMajorAxis
            txtInputCrsInvFlat.Text = CRS.Ellipsoid.InvFlattening
            txtInputCrsSemiMinAxis.Text = CRS.Ellipsoid.SemiMinorAxis
            txtInputCrsEllipsShape.Text = CRS.Ellipsoid.EllipsoidShape
            txtInputCrsEllipsoidRemarks.Text = CRS.Ellipsoid.Remarks

            'Prime meridian information:
            txtInputCrsPMName.Text = CRS.PrimeMeridian.Name
            txtInputCrsPMGreenLong.Text = CRS.PrimeMeridian.GreenwichLongitude
            txtInputCrsPMRemarks.Text = CRS.PrimeMeridian.Remarks
        Else
            DisplayInputBaseCrsLevelDatum(CRS.BaseCrs, Level, Count + 1)
        End If
    End Sub

    Private Sub DisplayOutputBaseCrsLevelDatum(CRS As CoordRefSystem, Level As Integer, Count As Integer)
        'Display the selected BaseCrs level
        If Count >= Level Then
            'Datum name:
            txtOutputCrsDatumName.Text = CRS.Datum.Name

            If Level = 0 Then
                GroupBox10.Text = "Datum:"
                Label76.Enabled = False
                txtOutputCrsBaseCrsName2.Enabled = False
                txtOutputCrsBaseCrsName2.Text = "" 'A BaseCrs Datum is not being displayed.
                Label74.Enabled = False
                txtOutputCrsBaseCrsCode2.Enabled = False
                txtOutputCrsBaseCrsCode2.Text = ""
                Label75.Enabled = False
                txtOutputCrsBaseCrsKind.Enabled = False
                txtOutputCrsBaseCrsKind.Text = ""
            Else
                GroupBox10.Text = "Base CRS Datum:"
                Label76.Enabled = True
                txtOutputCrsBaseCrsName2.Enabled = True
                txtOutputCrsBaseCrsName2.Text = CRS.Name
                Label74.Enabled = True
                txtOutputCrsBaseCrsCode2.Enabled = True
                txtOutputCrsBaseCrsCode2.Text = CRS.Code
                Label75.Enabled = True
                txtOutputCrsBaseCrsKind.Enabled = True
                Select Case CRS.Kind
                    Case CoordRefSystem.CrsKind.compound
                        txtOutputCrsBaseCrsKind.Text = "Compound"
                    Case CoordRefSystem.CrsKind.derived
                        txtOutputCrsBaseCrsKind.Text = "Derived"
                    Case CoordRefSystem.CrsKind.engineering
                        txtOutputCrsBaseCrsKind.Text = "Engineering"
                    Case CoordRefSystem.CrsKind.geocentric
                        txtOutputCrsBaseCrsKind.Text = "Geocentric"
                    Case CoordRefSystem.CrsKind.geodetic
                        txtOutputCrsBaseCrsKind.Text = "Geodetic"
                    Case CoordRefSystem.CrsKind.geographic2D
                        txtOutputCrsBaseCrsKind.Text = "Geographic 2D"
                    Case CoordRefSystem.CrsKind.geographic3D
                        txtOutputCrsBaseCrsKind.Text = "Geographic 3D"
                    Case CoordRefSystem.CrsKind.projected
                        txtOutputCrsBaseCrsKind.Text = "Projected"
                    Case CoordRefSystem.CrsKind.vertical
                        txtOutputCrsBaseCrsKind.Text = "Vertical"
                    Case Else
                        txtOutputCrsBaseCrsKind.Text = "Unrecognised: " & Conversion.OutputCrs.Kind.ToString
                End Select
            End If

            'Ellipsoid information:
            txtOutputCrsEllipsoidName.Text = CRS.Ellipsoid.Name
            txtOutputCrsSemiMajAxis.Text = CRS.Ellipsoid.SemiMajorAxis
            txtOutputCrsInvFlat.Text = CRS.Ellipsoid.InvFlattening
            txtOutputCrsSemiMinAxis.Text = CRS.Ellipsoid.SemiMinorAxis
            txtOutputCrsEllipsShape.Text = CRS.Ellipsoid.EllipsoidShape
            txtOutputCrsEllipsoidRemarks.Text = CRS.Ellipsoid.Remarks

            'Prime meridian information:
            txtOutputCrsPMName.Text = CRS.PrimeMeridian.Name
            txtOutputCrsPMGreenLong.Text = CRS.PrimeMeridian.GreenwichLongitude
            txtOutputCrsPMRemarks.Text = CRS.PrimeMeridian.Remarks
        Else
            DisplayOutputBaseCrsLevelDatum(CRS.BaseCrs, Level, Count + 1)
        End If
    End Sub


    Private Sub DisplayInputAxisInfo(AxisNo As Integer)
        'Display Axis information for the selected AxisNo

        Select Case AxisNo
            Case 0
                txtInputCrsCoordSysAxOrient.Text = ""
                txtInputCrsCoordSysAxAbbrev.Text = ""
                txtInputCrsCoordSysAxUom.Text = ""
                txtInputCrsCoordSysAxDescr.Text = ""
                txtInputCrsCoordSysAxRemarks.Text = ""
            Case 1
                If Conversion.InputCrs.CoordAxisList.Count > 0 Then
                    txtInputCrsCoordSysAxOrient.Text = Conversion.InputCrs.CoordAxisList(0).Orientation
                    txtInputCrsCoordSysAxAbbrev.Text = Conversion.InputCrs.CoordAxisList(0).Abbreviation
                    txtInputCrsCoordSysAxUom.Text = Conversion.UnitOfMeas(Conversion.InputCrs.CoordAxisList(0).UomCode).Name
                    txtInputCrsCoordSysAxDescr.Text = Conversion.InputCrs.CoordAxisNameList(0).Description
                    txtInputCrsCoordSysAxRemarks.Text = Conversion.InputCrs.CoordAxisNameList(0).Remarks
                Else
                    txtInputCrsCoordSysAxOrient.Text = ""
                    txtInputCrsCoordSysAxAbbrev.Text = ""
                    txtInputCrsCoordSysAxUom.Text = ""
                    txtInputCrsCoordSysAxDescr.Text = ""
                    txtInputCrsCoordSysAxRemarks.Text = ""
                End If
            Case 2
                If Conversion.InputCrs.CoordAxisList.Count > 1 Then
                    txtInputCrsCoordSysAxOrient.Text = Conversion.InputCrs.CoordAxisList(1).Orientation
                    txtInputCrsCoordSysAxAbbrev.Text = Conversion.InputCrs.CoordAxisList(1).Abbreviation
                    txtInputCrsCoordSysAxUom.Text = Conversion.UnitOfMeas(Conversion.InputCrs.CoordAxisList(1).UomCode).Name
                    txtInputCrsCoordSysAxDescr.Text = Conversion.InputCrs.CoordAxisNameList(1).Description
                    txtInputCrsCoordSysAxRemarks.Text = Conversion.InputCrs.CoordAxisNameList(1).Remarks
                Else
                    txtInputCrsCoordSysAxOrient.Text = ""
                    txtInputCrsCoordSysAxAbbrev.Text = ""
                    txtInputCrsCoordSysAxUom.Text = ""
                    txtInputCrsCoordSysAxDescr.Text = ""
                    txtInputCrsCoordSysAxRemarks.Text = ""
                End If

            Case 3
                If Conversion.InputCrs.CoordAxisList.Count > 2 Then
                    txtInputCrsCoordSysAxOrient.Text = Conversion.InputCrs.CoordAxisList(2).Orientation
                    txtInputCrsCoordSysAxAbbrev.Text = Conversion.InputCrs.CoordAxisList(2).Abbreviation
                    txtInputCrsCoordSysAxUom.Text = Conversion.UnitOfMeas(Conversion.InputCrs.CoordAxisList(2).UomCode).Name
                    txtInputCrsCoordSysAxDescr.Text = Conversion.InputCrs.CoordAxisNameList(2).Description
                    txtInputCrsCoordSysAxRemarks.Text = Conversion.InputCrs.CoordAxisNameList(2).Remarks
                Else
                    txtInputCrsCoordSysAxOrient.Text = ""
                    txtInputCrsCoordSysAxAbbrev.Text = ""
                    txtInputCrsCoordSysAxUom.Text = ""
                    txtInputCrsCoordSysAxDescr.Text = ""
                    txtInputCrsCoordSysAxRemarks.Text = ""
                End If
        End Select
    End Sub

    Private Sub DisplayOutputAxisInfo(AxisNo As Integer)
        'Display Axis information for the selected AxisNo

        Select Case AxisNo
            Case 0
                txtOutputCrsCoordSysAxOrient.Text = ""
                txtOutputCrsCoordSysAxAbbrev.Text = ""
                txtOutputCrsCoordSysAxUom.Text = ""
                txtOutputCrsCoordSysAxDescr.Text = ""
                txtOutputCrsCoordSysAxRemarks.Text = ""
            Case 1
                If Conversion.OutputCrs.CoordAxisList.Count > 0 Then
                    txtOutputCrsCoordSysAxOrient.Text = Conversion.OutputCrs.CoordAxisList(0).Orientation
                    txtOutputCrsCoordSysAxAbbrev.Text = Conversion.OutputCrs.CoordAxisList(0).Abbreviation
                    txtOutputCrsCoordSysAxUom.Text = Conversion.UnitOfMeas(Conversion.OutputCrs.CoordAxisList(0).UomCode).Name
                    txtOutputCrsCoordSysAxDescr.Text = Conversion.OutputCrs.CoordAxisNameList(0).Description
                    txtOutputCrsCoordSysAxRemarks.Text = Conversion.OutputCrs.CoordAxisNameList(0).Remarks
                Else
                    txtOutputCrsCoordSysAxOrient.Text = ""
                    txtOutputCrsCoordSysAxAbbrev.Text = ""
                    txtOutputCrsCoordSysAxUom.Text = ""
                    txtOutputCrsCoordSysAxDescr.Text = ""
                    txtOutputCrsCoordSysAxRemarks.Text = ""
                End If
            Case 2
                If Conversion.OutputCrs.CoordAxisList.Count > 1 Then
                    txtOutputCrsCoordSysAxOrient.Text = Conversion.OutputCrs.CoordAxisList(1).Orientation
                    txtOutputCrsCoordSysAxAbbrev.Text = Conversion.OutputCrs.CoordAxisList(1).Abbreviation
                    txtOutputCrsCoordSysAxUom.Text = Conversion.UnitOfMeas(Conversion.OutputCrs.CoordAxisList(1).UomCode).Name
                    txtOutputCrsCoordSysAxDescr.Text = Conversion.OutputCrs.CoordAxisNameList(1).Description
                    txtOutputCrsCoordSysAxRemarks.Text = Conversion.OutputCrs.CoordAxisNameList(1).Remarks
                Else
                    txtOutputCrsCoordSysAxOrient.Text = ""
                    txtOutputCrsCoordSysAxAbbrev.Text = ""
                    txtOutputCrsCoordSysAxUom.Text = ""
                    txtOutputCrsCoordSysAxDescr.Text = ""
                    txtOutputCrsCoordSysAxRemarks.Text = ""
                End If

            Case 3
                If Conversion.OutputCrs.CoordAxisList.Count > 2 Then
                    txtOutputCrsCoordSysAxOrient.Text = Conversion.OutputCrs.CoordAxisList(2).Orientation
                    txtOutputCrsCoordSysAxAbbrev.Text = Conversion.OutputCrs.CoordAxisList(2).Abbreviation
                    txtOutputCrsCoordSysAxUom.Text = Conversion.UnitOfMeas(Conversion.OutputCrs.CoordAxisList(2).UomCode).Name
                    txtOutputCrsCoordSysAxDescr.Text = Conversion.OutputCrs.CoordAxisNameList(2).Description
                    txtOutputCrsCoordSysAxRemarks.Text = Conversion.OutputCrs.CoordAxisNameList(2).Remarks
                Else
                    txtOutputCrsCoordSysAxOrient.Text = ""
                    txtOutputCrsCoordSysAxAbbrev.Text = ""
                    txtOutputCrsCoordSysAxUom.Text = ""
                    txtOutputCrsCoordSysAxDescr.Text = ""
                    txtOutputCrsCoordSysAxRemarks.Text = ""
                End If
        End Select
    End Sub

    Private Sub btnApplyInputQuery_Click(sender As Object, e As EventArgs) Handles btnApplyInputQuery.Click
        'Apply the CRS search query.
        ApplyInputCrsQuery()
    End Sub

    Private Sub btnApplyOutputQuery_Click(sender As Object, e As EventArgs) Handles btnApplyOutputQuery.Click
        'Apply the CRS search query.
        ApplyOutputCrsQuery()
    End Sub

    Private Sub ApplyInputCrsQuery()
        'Apply the Input CRS search query.

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

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(txtInputCrsQuery.Text, conn)
        If CrsInputSearch.Tables.Contains("List") Then CrsInputSearch.Tables("List").Clear() 'Clear any previous search results.
        da.Fill(CrsInputSearch, "List")
        udInputRowNo.Maximum = CrsInputSearch.Tables("List").Rows.Count - 1
        udInputRowNo.Value = -1
        conn.Close()

        dgvInputCrsList.DataSource = CrsInputSearch.Tables("List")
        'dgvCrsList.Update()
        'dgvCrsList.Refresh()
        txtNInputRecords.Text = CrsInputSearch.Tables("List").Rows.Count

        dgvInputCrsList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvInputCrsList.AutoResizeColumns()
    End Sub

    Private Sub ApplyOutputCrsQuery()
        'Apply the Output CRS search query.

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

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(txtOutputCrsQuery.Text, conn)
        If CrsOutputSearch.Tables.Contains("List") Then CrsOutputSearch.Tables("List").Clear() 'Clear any previous search results.
        da.Fill(CrsOutputSearch, "List")
        udOutputRowNo.Maximum = CrsOutputSearch.Tables("List").Rows.Count - 1
        udOutputRowNo.Value = -1
        conn.Close()

        dgvOutputCrsList.DataSource = CrsOutputSearch.Tables("List")
        'dgvCrsList.Update()
        'dgvCrsList.Refresh()
        txtNOutputRecords.Text = CrsOutputSearch.Tables("List").Rows.Count

        dgvOutputCrsList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvOutputCrsList.AutoResizeColumns()
    End Sub

    Private Sub btnFindInput_Click(sender As Object, e As EventArgs) Handles btnFindInput.Click

        If txtFindInput.Text.Trim = "" Then

        Else
            'txtInputCrsQuery.Text = "Select COORD_REF_SYS_CODE, COORD_REF_SYS_NAME, REMARKS From [Coordinate Reference System] Where COORD_REF_SYS_NAME Like '%" & txtFind.Text.Trim & "%'"
            txtInputCrsQuery.Text = "Select COORD_REF_SYS_CODE, COORD_REF_SYS_NAME, COORD_REF_SYS_KIND, REMARKS From [Coordinate Reference System] Where COORD_REF_SYS_NAME Like '%" & txtFindInput.Text.Trim & "%'"
            ApplyInputCrsQuery()
        End If
    End Sub

    Private Sub btnFindOutput_Click(sender As Object, e As EventArgs) Handles btnFindOutput.Click

        If txtFindOutput.Text.Trim = "" Then

        Else
            'txtOutputCrsQuery.Text = "Select COORD_REF_SYS_CODE, COORD_REF_SYS_NAME, REMARKS From [Coordinate Reference System] Where COORD_REF_SYS_NAME Like '%" & txtFind.Text.Trim & "%'"
            txtOutputCrsQuery.Text = "Select COORD_REF_SYS_CODE, COORD_REF_SYS_NAME, COORD_REF_SYS_KIND, REMARKS From [Coordinate Reference System] Where COORD_REF_SYS_NAME Like '%" & txtFindOutput.Text.Trim & "%'"
            ApplyOutputCrsQuery()
        End If
    End Sub

    Private Sub btnSelectInputCrs_Click(sender As Object, e As EventArgs) Handles btnSelectInputCrs.Click
        'Select the Input CRS
        Dim SelRowCount As Integer = dgvInputCrsList.SelectedRows.Count

        If SelRowCount = 0 Then

        ElseIf SelRowCount = 1 Then
            Dim SelRow As Integer = dgvInputCrsList.SelectedRows(0).Index
            Dim SelCrsCode As Integer = dgvInputCrsList.Rows(SelRow).Cells(0).Value
            udInputRowNo.Value = SelRow
            SelInputRowNo = SelRow
            'txtCrsCode.Text = SelCrsCode
            'Conversion.InputCRS.Code = Val(SelCrsCode)
            Conversion.InputCrs.Clear()
            Conversion.InputCrs.Code = SelCrsCode
            Conversion.InputCrs.GetAllSourceTargetCoordOps()
            ShowInputCrsInfo()
            DisplayDirectTransformationOptions()
            DisplayInputToWgs84TransOptions()
            DisplayWgs84ToOutputTransOptions()

        Else

        End If
    End Sub

    Private Sub btnSelectOutputCrs_Click(sender As Object, e As EventArgs) Handles btnSelectOutputCrs.Click
        'Select the Output CRS
        Dim SelRowCount As Integer = dgvOutputCrsList.SelectedRows.Count

        If SelRowCount = 0 Then

        ElseIf SelRowCount = 1 Then
            Dim SelRow As Integer = dgvOutputCrsList.SelectedRows(0).Index
            Dim SelCrsCode As Integer = dgvOutputCrsList.Rows(SelRow).Cells(0).Value
            udOutputRowNo.Value = SelRow
            SelOutputRowNo = SelRow
            'txtCrsCode.Text = SelCrsCode
            'Conversion.OutputCRS.Code = Val(SelCrsCode)
            Conversion.OutputCrs.Code = SelCrsCode
            Conversion.OutputCrs.GetAllSourceTargetCoordOps()
            ShowOutputCrsInfo()
            DisplayDirectTransformationOptions()
            DisplayInputToWgs84TransOptions()
            DisplayWgs84ToOutputTransOptions()

        Else

        End If
    End Sub

    Private Sub rbInputCrsCoordSysAx1Name_CheckedChanged(sender As Object, e As EventArgs) Handles rbInputCrsCoordSysAx1Name.CheckedChanged
        If rbInputCrsCoordSysAx1Name.Checked Then DisplayInputAxisInfo(1)
    End Sub

    Private Sub rbOutputCrsCoordSysAx1Name_CheckedChanged(sender As Object, e As EventArgs) Handles rbOutputCrsCoordSysAx1Name.CheckedChanged
        If rbOutputCrsCoordSysAx1Name.Checked Then DisplayOutputAxisInfo(1)
    End Sub

    Private Sub rbInputCrsCoordSysAx2Name_CheckedChanged(sender As Object, e As EventArgs) Handles rbInputCrsCoordSysAx2Name.CheckedChanged
        If rbInputCrsCoordSysAx2Name.Checked Then DisplayInputAxisInfo(2)
    End Sub

    Private Sub rbOutputCrsCoordSysAx2Name_CheckedChanged(sender As Object, e As EventArgs) Handles rbOutputCrsCoordSysAx2Name.CheckedChanged
        If rbOutputCrsCoordSysAx2Name.Checked Then DisplayOutputAxisInfo(2)
    End Sub

    Private Sub rbInputCrsCoordSysAx3Name_CheckedChanged(sender As Object, e As EventArgs) Handles rbInputCrsCoordSysAx3Name.CheckedChanged
        If rbInputCrsCoordSysAx3Name.Checked Then DisplayInputAxisInfo(3)
    End Sub

    Private Sub rbOutputCrsCoordSysAx3Name_CheckedChanged(sender As Object, e As EventArgs) Handles rbOutputCrsCoordSysAx3Name.CheckedChanged
        If rbOutputCrsCoordSysAx3Name.Checked Then DisplayOutputAxisInfo(3)
    End Sub

    Private Sub udInputFont_ValueChanged(sender As Object, e As EventArgs) Handles udInputFont.ValueChanged

        Dim NewFont As New Font(txtInputCrsProjMethodFormula.Font.Name, udInputFont.Value, txtInputCrsProjMethodFormula.Font.Style)
        txtInputCrsProjMethodFormula.Font = NewFont
        txtInputCrsProjMethodExample.Font = NewFont

    End Sub

    Private Sub udOutputFont_ValueChanged(sender As Object, e As EventArgs) Handles udOutputFont.ValueChanged

        Dim NewFont As New Font(txtOutputCrsProjMethodFormula.Font.Name, udOutputFont.Value, txtOutputCrsProjMethodFormula.Font.Style)
        txtOutputCrsProjMethodFormula.Font = NewFont
        txtOutputCrsProjMethodExample.Font = NewFont

    End Sub

    Private Sub udInputCalcFontSize_ValueChanged(sender As Object, e As EventArgs) Handles udInputCalcFontSize.ValueChanged

        Dim NewFont As New Font(dgvInputLocations.Columns(0).DefaultCellStyle.Font.Name, udInputCalcFontSize.Value, dgvInputLocations.Columns(0).DefaultCellStyle.Font.Style)
        dgvInputLocations.Columns(0).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(0).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(1).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(1).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(2).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(2).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(3).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(3).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(4).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(4).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(5).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(5).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(6).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(6).DefaultCellStyle.Font = NewFont
        dgvInputLocations.Columns(7).HeaderCell.Style.Font = NewFont
        dgvInputLocations.Columns(7).DefaultCellStyle.Font = NewFont
        dgvInputLocations.AutoResizeColumns()


        'If dgvInputLongLatToEastNorth.ColumnCount > 3 And dgvInputEastNorthToLongLat.ColumnCount > 3 Then
        '    'Dim NewFont As New Font(dgvLatLongToNorthEast.Columns(0).DefaultCellStyle.Font.Name, udInputCalcFontSize.Value, dgvLatLongToNorthEast.Columns(0).DefaultCellStyle.Font.Style)
        '    Dim NewFont As New Font(dgvInputLongLatToEastNorth.Columns(0).DefaultCellStyle.Font.Name, udInputCalcFontSize.Value, dgvInputLongLatToEastNorth.Columns(0).DefaultCellStyle.Font.Style)
        '    dgvInputLongLatToEastNorth.Columns(0).HeaderCell.Style.Font = NewFont
        '    dgvInputLongLatToEastNorth.Columns(0).DefaultCellStyle.Font = NewFont
        '    dgvInputLongLatToEastNorth.Columns(1).HeaderCell.Style.Font = NewFont
        '    dgvInputLongLatToEastNorth.Columns(1).DefaultCellStyle.Font = NewFont
        '    dgvInputLongLatToEastNorth.Columns(2).HeaderCell.Style.Font = NewFont
        '    dgvInputLongLatToEastNorth.Columns(2).DefaultCellStyle.Font = NewFont
        '    dgvInputLongLatToEastNorth.Columns(3).HeaderCell.Style.Font = NewFont
        '    dgvInputLongLatToEastNorth.Columns(3).DefaultCellStyle.Font = NewFont
        '    dgvInputLongLatToEastNorth.AutoResizeColumns()
        '    dgvInputLongLatToEastNorth.Width = dgvInputLongLatToEastNorth.RowHeadersWidth + dgvInputLongLatToEastNorth.Columns(0).Width + dgvInputLongLatToEastNorth.Columns(1).Width + dgvInputLongLatToEastNorth.Columns(2).Width + dgvInputLongLatToEastNorth.Columns(3).Width + 2

        '    dgvInputEastNorthToLongLat.Left = dgvInputLongLatToEastNorth.Left + dgvInputLongLatToEastNorth.Width + 4
        '    dgvInputEastNorthToLongLat.Columns(0).HeaderCell.Style.Font = NewFont
        '    dgvInputEastNorthToLongLat.Columns(0).DefaultCellStyle.Font = NewFont
        '    dgvInputEastNorthToLongLat.Columns(1).HeaderCell.Style.Font = NewFont
        '    dgvInputEastNorthToLongLat.Columns(1).DefaultCellStyle.Font = NewFont
        '    dgvInputEastNorthToLongLat.Columns(2).HeaderCell.Style.Font = NewFont
        '    dgvInputEastNorthToLongLat.Columns(2).DefaultCellStyle.Font = NewFont
        '    dgvInputEastNorthToLongLat.Columns(3).HeaderCell.Style.Font = NewFont
        '    dgvInputEastNorthToLongLat.Columns(3).DefaultCellStyle.Font = NewFont
        '    dgvInputEastNorthToLongLat.AutoResizeColumns()
        '    dgvInputEastNorthToLongLat.Width = dgvInputEastNorthToLongLat.RowHeadersWidth + dgvInputEastNorthToLongLat.Columns(0).Width + dgvInputEastNorthToLongLat.Columns(1).Width + dgvInputEastNorthToLongLat.Columns(2).Width + dgvInputEastNorthToLongLat.Columns(3).Width + 2
        'End If
    End Sub

    Private Sub udOutputCalcFontSize_ValueChanged(sender As Object, e As EventArgs) Handles udOutputCalcFontSize.ValueChanged

        Dim NewFont As New Font(dgvOutputLocations.Columns(0).DefaultCellStyle.Font.Name, udOutputCalcFontSize.Value, dgvOutputLocations.Columns(0).DefaultCellStyle.Font.Style)
        dgvOutputLocations.Columns(0).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(0).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(1).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(1).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(2).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(2).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(3).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(3).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(4).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(4).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(5).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(5).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(6).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(6).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.Columns(7).HeaderCell.Style.Font = NewFont
        dgvOutputLocations.Columns(7).DefaultCellStyle.Font = NewFont
        dgvOutputLocations.AutoResizeColumns()

        'If dgvOutputLongLatToEastNorth.ColumnCount > 3 And dgvOutputEastNorthToLongLat.ColumnCount > 3 Then
        '    'Dim NewFont As New Font(dgvLatLongToNorthEast.Columns(0).DefaultCellStyle.Font.Name, udOutputCalcFontSize.Value, dgvLatLongToNorthEast.Columns(0).DefaultCellStyle.Font.Style)
        '    Dim NewFont As New Font(dgvOutputLongLatToEastNorth.Columns(0).DefaultCellStyle.Font.Name, udOutputCalcFontSize.Value, dgvOutputLongLatToEastNorth.Columns(0).DefaultCellStyle.Font.Style)
        '    dgvOutputLongLatToEastNorth.Columns(0).HeaderCell.Style.Font = NewFont
        '    dgvOutputLongLatToEastNorth.Columns(0).DefaultCellStyle.Font = NewFont
        '    dgvOutputLongLatToEastNorth.Columns(1).HeaderCell.Style.Font = NewFont
        '    dgvOutputLongLatToEastNorth.Columns(1).DefaultCellStyle.Font = NewFont
        '    dgvOutputLongLatToEastNorth.Columns(2).HeaderCell.Style.Font = NewFont
        '    dgvOutputLongLatToEastNorth.Columns(2).DefaultCellStyle.Font = NewFont
        '    dgvOutputLongLatToEastNorth.Columns(3).HeaderCell.Style.Font = NewFont
        '    dgvOutputLongLatToEastNorth.Columns(3).DefaultCellStyle.Font = NewFont
        '    dgvOutputLongLatToEastNorth.AutoResizeColumns()
        '    dgvOutputLongLatToEastNorth.Width = dgvOutputLongLatToEastNorth.RowHeadersWidth + dgvOutputLongLatToEastNorth.Columns(0).Width + dgvOutputLongLatToEastNorth.Columns(1).Width + dgvOutputLongLatToEastNorth.Columns(2).Width + dgvOutputLongLatToEastNorth.Columns(3).Width + 2

        '    dgvOutputEastNorthToLongLat.Left = dgvOutputLongLatToEastNorth.Left + dgvOutputLongLatToEastNorth.Width + 4
        '    dgvOutputEastNorthToLongLat.Columns(0).HeaderCell.Style.Font = NewFont
        '    dgvOutputEastNorthToLongLat.Columns(0).DefaultCellStyle.Font = NewFont
        '    dgvOutputEastNorthToLongLat.Columns(1).HeaderCell.Style.Font = NewFont
        '    dgvOutputEastNorthToLongLat.Columns(1).DefaultCellStyle.Font = NewFont
        '    dgvOutputEastNorthToLongLat.Columns(2).HeaderCell.Style.Font = NewFont
        '    dgvOutputEastNorthToLongLat.Columns(2).DefaultCellStyle.Font = NewFont
        '    dgvOutputEastNorthToLongLat.Columns(3).HeaderCell.Style.Font = NewFont
        '    dgvOutputEastNorthToLongLat.Columns(3).DefaultCellStyle.Font = NewFont
        '    dgvOutputEastNorthToLongLat.AutoResizeColumns()
        '    dgvOutputEastNorthToLongLat.Width = dgvOutputEastNorthToLongLat.RowHeadersWidth + dgvOutputEastNorthToLongLat.Columns(0).Width + dgvOutputEastNorthToLongLat.Columns(1).Width + dgvOutputEastNorthToLongLat.Columns(2).Width + dgvOutputEastNorthToLongLat.Columns(3).Width + 2
        'End If
    End Sub

    'Private Sub ResizeInputCalcGrids()
    '    'Resize the Input Calculations Data Grid Views.
    '    dgvInputLongLatToEastNorth.AutoResizeColumns()
    '    dgvInputEastNorthToLongLat.AutoResizeColumns()
    '    dgvInputLongLatToEastNorth.Width = dgvInputLongLatToEastNorth.RowHeadersWidth + dgvInputLongLatToEastNorth.Columns(0).Width + dgvInputLongLatToEastNorth.Columns(1).Width + dgvInputLongLatToEastNorth.Columns(2).Width + dgvInputLongLatToEastNorth.Columns(3).Width + 2
    '    dgvInputEastNorthToLongLat.Left = dgvInputLongLatToEastNorth.Left + dgvInputLongLatToEastNorth.Width + 4
    '    dgvInputEastNorthToLongLat.Width = dgvInputEastNorthToLongLat.RowHeadersWidth + dgvInputEastNorthToLongLat.Columns(0).Width + dgvInputEastNorthToLongLat.Columns(1).Width + dgvInputEastNorthToLongLat.Columns(2).Width + dgvInputEastNorthToLongLat.Columns(3).Width + 2
    'End Sub

    'Private Sub ResizeOutputCalcGrids()
    '    'Resize the Output Calculations Data Grid Views.
    '    dgvOutputLongLatToEastNorth.AutoResizeColumns()
    '    dgvOutputLongLatToEastNorth.Width = dgvOutputLongLatToEastNorth.RowHeadersWidth + dgvOutputLongLatToEastNorth.Columns(0).Width + dgvOutputLongLatToEastNorth.Columns(1).Width + dgvOutputLongLatToEastNorth.Columns(2).Width + dgvOutputLongLatToEastNorth.Columns(3).Width + 2
    '    dgvOutputEastNorthToLongLat.Left = dgvOutputLongLatToEastNorth.Left + dgvOutputLongLatToEastNorth.Width + 4
    '    dgvOutputEastNorthToLongLat.Width = dgvOutputEastNorthToLongLat.RowHeadersWidth + dgvOutputEastNorthToLongLat.Columns(0).Width + dgvOutputEastNorthToLongLat.Columns(1).Width + dgvOutputEastNorthToLongLat.Columns(2).Width + dgvOutputEastNorthToLongLat.Columns(3).Width + 2
    'End Sub

    Private Sub btnOpenInputCrsCode_Click(sender As Object, e As EventArgs) Handles btnOpenInputCrsCode.Click
        'Open the CRS with the code shown in txtInputCrsCode
        Conversion.InputCrs.Code = Val(txtInputCrsCode.Text)
        Conversion.InputCrs.GetAllSourceTargetCoordOps()
        ShowInputCrsInfo()
    End Sub

    Private Sub btnOpenOutputCrsCode_Click(sender As Object, e As EventArgs) Handles btnOpenOutputCrsCode.Click
        'Open the CRS with the code shown in txtOutputCrsCode
        Conversion.OutputCrs.Code = Val(txtOutputCrsCode.Text)
        Conversion.OutputCrs.GetAllSourceTargetCoordOps()
        ShowOutputCrsInfo()
    End Sub

    Private Sub udInputRowNo_ValueChanged(sender As Object, e As EventArgs) Handles udInputRowNo.ValueChanged

        If udInputRowNo.Focused Then
            'Only select the new InputCRS if the control has hte focus.
            SelInputRowNo = udInputRowNo.Value
            If SelInputRowNo = -1 Then

            Else
                If dgvInputCrsList.Rows.Count > SelInputRowNo + 1 Then
                    'SelMethodCode = dgvCrsList.Rows(SelInputRowNo).Cells(0).Value
                    Conversion.InputCrs.Code = dgvInputCrsList.Rows(SelInputRowNo).Cells(0).Value
                    Conversion.InputCrs.GetAllSourceTargetCoordOps()
                    ShowInputCrsInfo()
                    dgvInputCrsList.ClearSelection()
                    dgvInputCrsList.Rows(SelInputRowNo).Selected = True
                End If
            End If
        End If
    End Sub

    Private Sub udOutputRowNo_ValueChanged(sender As Object, e As EventArgs) Handles udOutputRowNo.ValueChanged

        If udOutputRowNo.Focused Then
            'Only select the new OutputCRS if the control has hte focus.
            SelOutputRowNo = udOutputRowNo.Value
            If SelOutputRowNo = -1 Then

            Else
                If dgvOutputCrsList.Rows.Count > SelOutputRowNo + 1 Then
                    'SelMethodCode = dgvCrsList.Rows(SelOutputRowNo).Cells(0).Value
                    Conversion.OutputCrs.Code = dgvOutputCrsList.Rows(SelOutputRowNo).Cells(0).Value
                    Conversion.OutputCrs.GetAllSourceTargetCoordOps()
                    ShowOutputCrsInfo()
                    dgvOutputCrsList.ClearSelection()
                    dgvOutputCrsList.Rows(SelOutputRowNo).Selected = True
                End If
            End If
        End If
    End Sub

    Private Sub dgvLongLatToEastNorth_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    'Private Sub dgvInputLongLatToEastNorth_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
    '    'Cell End Edit - update the Northing Easting calculation
    '    Dim ColNo As Integer = e.ColumnIndex
    '    'If ColNo = 0 Then 'Latitude value changed.
    '    If ColNo = 0 Then 'Longitude value changed.
    '        Dim RowNo As Integer = e.RowIndex
    '        If dgvInputLongLatToEastNorth.Rows(RowNo).Cells(1).Value = "" Then
    '            'No Latitude value.
    '        Else
    '            'Conversion.InputCrs.Projection.Method.Latitude = Val(dgvInputLatLongToNorthEast.Rows(RowNo).Cells(0).Value)
    '            'Conversion.InputCrs.Projection.Method.Coord.Latitude = Val(dgvInputLongLatToEastNorth.Rows(RowNo).Cells(0).Value)
    '            Conversion.InputCrs.Projection.Method.Coord.Longitude = Val(dgvInputLongLatToEastNorth.Rows(RowNo).Cells(0).Value)
    '            'Conversion.InputCrs.Projection.Method.Longitude = Val(dgvInputLatLongToNorthEast.Rows(RowNo).Cells(1).Value)
    '            'Conversion.InputCrs.Projection.Method.Coord.Longitude = Val(dgvInputLongLatToEastNorth.Rows(RowNo).Cells(1).Value)
    '            Conversion.InputCrs.Projection.Method.Coord.Latitude = Val(dgvInputLongLatToEastNorth.Rows(RowNo).Cells(1).Value)
    '            Conversion.InputCrs.Projection.Method.LongLatToEastNorth
    '            'dgvInputLatLongToNorthEast.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Northing
    '            'dgvInputLongLatToEastNorth.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Coord.Northing
    '            dgvInputLongLatToEastNorth.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Coord.Easting
    '            'dgvInputLatLongToNorthEast.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Easting
    '            'dgvInputLongLatToEastNorth.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Coord.Easting
    '            dgvInputLongLatToEastNorth.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Coord.Northing
    '            ResizeInputCalcGrids()
    '        End If
    '        'ElseIf ColNo = 1 Then 'Longitude value changed.
    '    ElseIf ColNo = 1 Then 'Latitude value changed.
    '        Dim RowNo As Integer = e.RowIndex
    '        If dgvInputLongLatToEastNorth.Rows(RowNo).Cells(0).Value = "" Then
    '            'No Longitude value.
    '        Else
    '            'Conversion.InputCrs.Projection.Method.Latitude = Val(dgvInputLatLongToNorthEast.Rows(RowNo).Cells(0).Value)
    '            'Conversion.InputCrs.Projection.Method.Coord.Latitude = Val(dgvInputLongLatToEastNorth.Rows(RowNo).Cells(0).Value)
    '            Conversion.InputCrs.Projection.Method.Coord.Longitude = Val(dgvInputLongLatToEastNorth.Rows(RowNo).Cells(0).Value)
    '            'Conversion.InputCrs.Projection.Method.Longitude = Val(dgvInputLatLongToNorthEast.Rows(RowNo).Cells(1).Value)
    '            'Conversion.InputCrs.Projection.Method.Coord.Longitude = Val(dgvInputLongLatToEastNorth.Rows(RowNo).Cells(1).Value)
    '            Conversion.InputCrs.Projection.Method.Coord.Latitude = Val(dgvInputLongLatToEastNorth.Rows(RowNo).Cells(1).Value)
    '            Conversion.InputCrs.Projection.Method.LongLatToEastNorth
    '            'dgvInputLatLongToNorthEast.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Northing
    '            'dgvInputLongLatToEastNorth.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Coord.Northing
    '            dgvInputLongLatToEastNorth.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Coord.Easting
    '            'dgvInputLatLongToNorthEast.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Easting
    '            'dgvInputLongLatToEastNorth.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Coord.Easting
    '            dgvInputLongLatToEastNorth.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Coord.Northing
    '            ResizeInputCalcGrids()
    '        End If
    '    Else

    '    End If

    'End Sub

    'Private Sub dgvOutputLongLatToEastNorth_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
    '    'Cell End Edit - update the Northing Easting calculation
    '    Dim ColNo As Integer = e.ColumnIndex
    '    'If ColNo = 0 Then 'Latitude value changed.
    '    If ColNo = 0 Then 'Longitude value changed.
    '        Dim RowNo As Integer = e.RowIndex
    '        If dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(1).Value = "" Then
    '            'No Longitude value.
    '        Else
    '            'Conversion.OutputCrs.Projection.Method.Latitude = Val(dgvOutputLatLongToNorthEast.Rows(RowNo).Cells(0).Value)
    '            'Conversion.OutputCrs.Projection.Method.Coord.Latitude = Val(dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(0).Value)
    '            Conversion.OutputCrs.Projection.Method.Coord.Longitude = Val(dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(0).Value)
    '            'Conversion.OutputCrs.Projection.Method.Longitude = Val(dgvOutputLatLongToNorthEast.Rows(RowNo).Cells(1).Value)
    '            'Conversion.OutputCrs.Projection.Method.Coord.Longitude = Val(dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(1).Value)
    '            Conversion.OutputCrs.Projection.Method.Coord.Latitude = Val(dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(1).Value)
    '            Conversion.OutputCrs.Projection.Method.LongLatToEastNorth
    '            'dgvOutputLatLongToNorthEast.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Northing
    '            'dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Coord.Northing
    '            dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Coord.Easting
    '            'dgvOutputLatLongToNorthEast.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Easting
    '            'dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Coord.Easting
    '            dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Coord.Northing
    '            ResizeOutputCalcGrids()
    '        End If
    '        'ElseIf ColNo = 1 Then 'Longitude value changed.
    '    ElseIf ColNo = 1 Then 'Latitude value changed.
    '        Dim RowNo As Integer = e.RowIndex
    '        If dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(0).Value = "" Then
    '            'No Longitude value.
    '        Else
    '            'Conversion.OutputCrs.Projection.Method.Latitude = Val(dgvOutputLatLongToNorthEast.Rows(RowNo).Cells(0).Value)
    '            'Conversion.OutputCrs.Projection.Method.Coord.Latitude = Val(dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(0).Value)
    '            Conversion.OutputCrs.Projection.Method.Coord.Longitude = Val(dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(0).Value)
    '            'Conversion.OutputCrs.Projection.Method.Longitude = Val(dgvOutputLatLongToNorthEast.Rows(RowNo).Cells(1).Value)
    '            'Conversion.OutputCrs.Projection.Method.Coord.Longitude = Val(dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(1).Value)
    '            Conversion.OutputCrs.Projection.Method.Coord.Latitude = Val(dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(1).Value)
    '            Conversion.OutputCrs.Projection.Method.LongLatToEastNorth
    '            'dgvOutputLatLongToNorthEast.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Northing
    '            'dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Coord.Northing
    '            dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Coord.Easting
    '            'dgvOutputLatLongToNorthEast.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Easting
    '            'dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Coord.Easting
    '            dgvOutputLongLatToEastNorth.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Coord.Northing
    '            ResizeOutputCalcGrids()
    '        End If
    '    Else

    '    End If

    'End Sub

    Private Sub dgvEastNorthToLongLat_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    'Private Sub dgvInputEastNorthToLongLat_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
    '    'Cell End Edit - update the Latitude Longitude calculation
    '    Dim ColNo As Integer = e.ColumnIndex
    '    'If ColNo = 0 Then 'Northing value changed.
    '    If ColNo = 0 Then 'Easting value changed.
    '        Dim RowNo As Integer = e.RowIndex
    '        If dgvInputEastNorthToLongLat.Rows(RowNo).Cells(1).Value = "" Then
    '            'No Northing value.
    '        Else
    '            'Conversion.InputCrs.Projection.Method.Northing = Val(dgvInputNorthEastToLatLong.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            'Conversion.InputCrs.Projection.Method.Coord.Northing = Val(dgvInputEastNorthToLongLat.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            Conversion.InputCrs.Projection.Method.Coord.Easting = Val(dgvInputEastNorthToLongLat.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            'Conversion.InputCrs.Projection.Method.Easting = Val(dgvInputNorthEastToLatLong.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            'Conversion.InputCrs.Projection.Method.Coord.Easting = Val(dgvInputEastNorthToLongLat.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            Conversion.InputCrs.Projection.Method.Coord.Northing = Val(dgvInputEastNorthToLongLat.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            Conversion.InputCrs.Projection.Method.EastNorthToLongLat
    '            'dgvInputNorthEastToLatLong.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Latitude
    '            'dgvInputEastNorthToLongLat.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Coord.Latitude
    '            dgvInputEastNorthToLongLat.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Coord.Longitude
    '            'dgvInputNorthEastToLatLong.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Longitude
    '            'dgvInputEastNorthToLongLat.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Coord.Longitude
    '            dgvInputEastNorthToLongLat.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Coord.Latitude
    '            ResizeInputCalcGrids()
    '        End If
    '    ElseIf ColNo = 1 Then 'Northing value changed.
    '        Dim RowNo As Integer = e.RowIndex
    '        If dgvInputEastNorthToLongLat.Rows(RowNo).Cells(0).Value = "" Then
    '            'No Easting value.
    '        Else
    '            'Conversion.InputCrs.Projection.Method.Northing = Val(dgvInputNorthEastToLatLong.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            'Conversion.InputCrs.Projection.Method.Coord.Northing = Val(dgvInputEastNorthToLongLat.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            Conversion.InputCrs.Projection.Method.Coord.Easting = Val(dgvInputEastNorthToLongLat.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            'Conversion.InputCrs.Projection.Method.Easting = Val(dgvInputNorthEastToLatLong.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            'Conversion.InputCrs.Projection.Method.Coord.Easting = Val(dgvInputEastNorthToLongLat.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            Conversion.InputCrs.Projection.Method.Coord.Northing = Val(dgvInputEastNorthToLongLat.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            Conversion.InputCrs.Projection.Method.EastNorthToLongLat
    '            'dgvInputNorthEastToLatLong.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Latitude
    '            'dgvInputEastNorthToLongLat.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Coord.Latitude
    '            dgvInputEastNorthToLongLat.Rows(RowNo).Cells(2).Value = Conversion.InputCrs.Projection.Method.Coord.Longitude
    '            'dgvInputNorthEastToLatLong.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Longitude
    '            'dgvInputEastNorthToLongLat.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Coord.Longitude
    '            dgvInputEastNorthToLongLat.Rows(RowNo).Cells(3).Value = Conversion.InputCrs.Projection.Method.Coord.Latitude
    '            ResizeInputCalcGrids()
    '        End If
    '    Else

    '    End If
    'End Sub

    'Private Sub dgvOutputEastNorthToLongLat_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
    '    'Cell End Edit - update the Latitude Longitude calculation
    '    Dim ColNo As Integer = e.ColumnIndex
    '    'If ColNo = 0 Then 'Northing value changed.
    '    If ColNo = 0 Then 'Easting value changed.
    '        Dim RowNo As Integer = e.RowIndex
    '        If dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(1).Value = "" Then
    '            'No Northing value.
    '        Else
    '            'Conversion.OutputCrs.Projection.Method.Northing = Val(dgvOutputNorthEastToLatLong.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            'Conversion.OutputCrs.Projection.Method.Coord.Northing = Val(dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            Conversion.OutputCrs.Projection.Method.Coord.Easting = Val(dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            'Conversion.OutputCrs.Projection.Method.Easting = Val(dgvOutputNorthEastToLatLong.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            'Conversion.OutputCrs.Projection.Method.Coord.Easting = Val(dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            Conversion.OutputCrs.Projection.Method.Coord.Northing = Val(dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            Conversion.OutputCrs.Projection.Method.EastNorthToLongLat
    '            'dgvOutputNorthEastToLatLong.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Latitude
    '            'dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Coord.Latitude
    '            dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Coord.Longitude
    '            'dgvOutputNorthEastToLatLong.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Longitude
    '            'dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Coord.Longitude
    '            dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Coord.Latitude
    '            ResizeOutputCalcGrids()
    '        End If
    '        'ElseIf ColNo = 1 Then 'Easting value changed.
    '    ElseIf ColNo = 1 Then 'Northing value changed.
    '        Dim RowNo As Integer = e.RowIndex
    '        If dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(0).Value = "" Then
    '            'No Easting value.
    '        Else
    '            'Conversion.OutputCrs.Projection.Method.Northing = Val(dgvOutputNorthEastToLatLong.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            'Conversion.OutputCrs.Projection.Method.Coord.Northing = Val(dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            Conversion.OutputCrs.Projection.Method.Coord.Easting = Val(dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(0).Value.Replace(",", ""))
    '            'Conversion.OutputCrs.Projection.Method.Easting = Val(dgvOutputNorthEastToLatLong.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            'Conversion.OutputCrs.Projection.Method.Coord.Easting = Val(dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            Conversion.OutputCrs.Projection.Method.Coord.Northing = Val(dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(1).Value.Replace(",", ""))
    '            Conversion.OutputCrs.Projection.Method.EastNorthToLongLat
    '            'dgvOutputNorthEastToLatLong.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Latitude
    '            'dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Coord.Latitude
    '            dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(2).Value = Conversion.OutputCrs.Projection.Method.Coord.Longitude
    '            'dgvOutputNorthEastToLatLong.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Longitude
    '            'dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Coord.Longitude
    '            dgvOutputEastNorthToLongLat.Rows(RowNo).Cells(3).Value = Conversion.OutputCrs.Projection.Method.Coord.Latitude
    '            ResizeOutputCalcGrids()
    '        End If
    '    Else

    '    End If
    'End Sub

    Private Sub txtInputCrsProjMethodFormula_TextChanged(sender As Object, e As EventArgs) Handles txtInputCrsProjMethodFormula.TextChanged

    End Sub

    Private Sub txtInputCrsProjMethodFormula_MouseHover(sender As Object, e As EventArgs) Handles txtInputCrsProjMethodFormula.MouseHover

    End Sub

    Private Sub txtInputCrsProjMethodFormula_MouseMove(sender As Object, e As MouseEventArgs) Handles txtInputCrsProjMethodFormula.MouseMove

    End Sub

    Private Sub txtInputCrsProjMethodFormula_MouseUp(sender As Object, e As MouseEventArgs) Handles txtInputCrsProjMethodFormula.MouseUp
        'Show character information for the first selected character.
        If txtInputCrsProjMethodFormula.SelectedText.Length > 0 Then
            'txtInputCharCode.Text = Asc(txtInputCrsProjMethodFormula.SelectedText.First)
            Dim CharCode As Integer = AscW(txtInputCrsProjMethodFormula.SelectedText)
            'txtInputCharCode.Text = AscW(txtInputCrsProjMethodFormula.SelectedText)
            'txtInputCharCode.Text = CharCode
            If chkInputHex.Checked Then txtInputCharCode.Text = Hex(CharCode) Else txtInputCharCode.Text = CharCode
            txtInputCharName.Text = CharName(CharCode)
        End If
    End Sub

    Private Sub txtOutputCrsProjMethodFormula_MouseUp(sender As Object, e As MouseEventArgs) Handles txtOutputCrsProjMethodFormula.MouseUp
        'Show character information for the first selected character.
        If txtOutputCrsProjMethodFormula.SelectedText.Length > 0 Then
            Dim CharCode As Integer = AscW(txtOutputCrsProjMethodFormula.SelectedText)
            If chkOutputHex.Checked Then txtOutputCharCode.Text = Hex(CharCode) Else txtOutputCharCode.Text = CharCode
            txtOutputCharName.Text = CharName(CharCode)
        End If
    End Sub

    Private Function CharName(CharCode As Integer) As String
        'Return the character name corresponding to a character code.

        If CharCode <= 31 Then 'Control character
            Return "Control character"
        ElseIf CharCode <= 47 Then 'Printable character
            Return "Printable character"
        ElseIf CharCode <= 57 Then 'digit character
            Return "Digit"
        ElseIf CharCode <= 64 Then 'Printable character
            Return "Printable character"
        ElseIf CharCode <= 90 Then 'Upper chase letter
            Return "Upper case letter"
        ElseIf CharCode <= 96 Then 'Printable character
            Return "Printable character"
        ElseIf CharCode <= 122 Then 'Lower case letter
            Return "Lower case letter"
        ElseIf CharCode <= 127 Then 'Printable character
            Return "Printable character"
        ElseIf CharCode <= 255 Then 'Extended ASCII code
            Return "Extended ASCII character"
        ElseIf CharCode <= 912 Then
            Return "Extended ASCII character"
        ElseIf CharCode <= 937 Then 'Greek capital letter
            Return GreekCapitalLetterName(CharCode)
        ElseIf CharCode <= 969 Then 'Greek small letter
            Return GreekLowerCaseLetterName(CharCode)
        Else
            Return "Extended ASCII character"
        End If
    End Function

    Function GreekCapitalLetterName(CharCode As Integer) As String
        'Return the Control character name corresponding to a Character Code.
        Select Case CharCode
            Case 913
                Return "Greek capital letter Alpha"
            Case 914
                Return "Greek capital letter Beta"
            Case 915
                Return "Greek capital letter Gamma"
            Case 916
                Return "Greek capital letter Delta"
            Case 917
                Return "Greek capital letter Epsilon"
            Case 918
                Return "Greek capital letter Zeta"
            Case 919
                Return "Greek capital letter Eta"
            Case 920
                Return "Greek capital letter Theta"
            Case 921
                Return "Greek capital letter Iota"
            Case 922
                Return "Greek capital letter Kappa"
            Case 923
                Return "Greek capital letter Lambda"
            Case 924
                Return "Greek capital letter Mu"
            Case 925
                Return "Greek capital letter Nu"
            Case 926
                Return "Greek capital letter Xi"
            Case 927
                Return "Greek capital letter Omicron"
            Case 928
                Return "Greek capital letter Pi"
            Case 929
                Return "Greek capital letter Rho"
            Case 930
                Return ""
            Case 931
                Return "Greek capital letter Sigma"
            Case 932
                Return "Greek capital letter Tau"
            Case 933
                Return "Greek capital letter Upsilon"
            Case 934
                Return "Greek capital letter Phi"
            Case 935
                Return "Greek capital letter Chi"
            Case 936
                Return "Greek capital letter Psi"
            Case 937
                Return "Greek capital letter Omega"
            Case Else
                Return ""
        End Select
    End Function

    Function GreekLowerCaseLetterName(CharCode As Integer) As String
        'Return the Control character name corresponding to a Character Code.
        Select Case CharCode
            Case 945
                Return "Greek lower case letter Alpha"
            Case 946
                Return "Greek lower case letter Beta"
            Case 947
                Return "Greek lower case letter Gamme"
            Case 948
                Return "Greek lower case letter Delta"
            Case 949
                Return "Greek lower case letter Epsilon"
            Case 950
                Return "Greek lower case letter Zeta"
            Case 951
                Return "Greek lower case letter Eta"
            Case 952
                Return "Greek lower case letter Theta"
            Case 953
                Return "Greek lower case letter Iota"
            Case 954
                Return "Greek lower case letter Kappa"
            Case 955
                Return "Greek lower case letter Lambda"
            Case 956
                Return "Greek lower case letter Mu"
            Case 957
                Return "Greek lower case letter Nu"
            Case 958
                Return "Greek lower case letter Xi"
            Case 959
                Return "Greek lower case letter Omicron"
            Case 960
                Return "Greek lower case letter Pi"
            Case 961
                Return "Greek lower case letter Rho"
            Case 962
                Return "Greek lower case letter Final Sigma"
            Case 963
                Return "Greek lower case letter Sigma"
            Case 964
                Return "Greek lower case letter Tau"
            Case 965
                Return "Greek lower case letter Upsilon"
            Case 966
                Return "Greek lower case letter Phi"
            Case 967
                Return "Greek lower case letter Chi"
            Case 968
                Return "Greek lower case letter Psi"
            Case 969
                Return "Greek lower case letter Omega"
            Case Else
                Return ""
        End Select
    End Function

    Private Sub btnFormatHelp_Click(sender As Object, e As EventArgs) Handles btnFormatHelp.Click
        'Show Format information.
        MessageBox.Show("Format string examples:" & vbCrLf & "N4 - Number displayed with thousands separator and 4 decimal places" & vbCrLf & "F4 - Number displayed with 4 decimal places.", "Number Formatting")
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl2.SelectedIndexChanged
        'Restore the Splitter Distance for the Input Coordinate Operation Method.
        If TabControl2.SelectedIndex = 2 Then 'Coordinate Operation Method tab selected.
            SplitContainer1.SplitterDistance = SplitDist1
            SplitContainer2.SplitterDistance = SplitDist2
        ElseIf TabControl2.SelectedIndex = 4 Then 'Coordinate Operations tab selected.
            SplitContainer5.SplitterDistance = SplitDist5
        End If
    End Sub

    Private Sub TabControl3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl3.SelectedIndexChanged
        'Restore the Splitter Distance for the Output Coordinate Operation Method.
        If TabControl3.SelectedIndex = 2 Then 'Coordinate Operation Method tab selected.
            SplitContainer3.SplitterDistance = SplitDist3
            SplitContainer4.SplitterDistance = SplitDist4
        ElseIf TabControl2.SelectedIndex = 4 Then 'Coordinate Operations tab selected.
            SplitContainer6.SplitterDistance = SplitDist6
        End If
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles udOutputDatumLevel.ValueChanged

    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles txtOutputCrsBaseCrsKind.TextChanged

    End Sub

    Private Sub TextBox36_TextChanged(sender As Object, e As EventArgs) Handles txtOutputCrsDatumName.TextChanged

    End Sub

    Private Sub TextBox33_TextChanged(sender As Object, e As EventArgs) Handles txtOutputCrsInvFlat.TextChanged

    End Sub

    Private Sub dgvOutputCrsList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOutputCrsList.CellContentClick

    End Sub

    Private Sub dgvInputCrsList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvInputCrsList.CellContentClick

    End Sub



    Private Sub DisplayTransformationOptions(InputCrs As CoordRefSystem, OutputCrs As CoordRefSystem)
        'Display the Transformation Options on the Datum Transformation Operation tab.

        'Clear the Default Datum Transformation information:
        txtDefDatumTransType.Text = ""
        txtDefDatumTransOpName1.Text = ""
        txtDefDatumTransOpCode1.Text = ""
        txtDefDatumTransOpName2.Text = ""
        txtDefDatumTransOpCode2.Text = ""

        If InputCrs.DatumCode = -1 Then
            If IsNothing(InputCrs.BaseCrs) Then
                'The Input CRS has an unknown datum.

            Else
                DisplayTransformationOptions(InputCrs.BaseCrs, OutputCrs)
            End If
        ElseIf OutputCrs.DatumCode = -1 Then
            If IsNothing(OutputCrs.BaseCrs) Then
                'The Output CRS has an unknown datum.

            Else
                DisplayTransformationOptions(InputCrs, OutputCrs.BaseCrs)
            End If
        Else
            If InputCrs.DatumCode = OutputCrs.DatumCode Then
                'Datum transformation not required.
                rbDtNotRequired.Enabled = True
                rbDtNotRequired.Checked = True
            ElseIf Conversion.DirectDatumTransOpList.Count > 0 Then
                'Direct Datum Transformation operations are available
                rbDtNotRequired.Enabled = False
                rbDirectDatumTrans.Checked = True
                'FindDefaultDatumTrans()
            Else
                rbDtNotRequired.Enabled = False
                rbDatumTransViaWgs84.Checked = True
                'FindDefaultDatumTrans()
            End If
        End If

    End Sub

    Private Sub DisplayDirectTransformationOptions()
        'Display the coordinate operations that could be used for a direct datum transformation from the Input Crs to the Output Crs.
        If Conversion.InputCrs.Code > -1 And Conversion.OutputCrs.Code > -1 Then
            dgvDirectDTOps.Rows.Clear()
            For Each Item In Conversion.DirectDatumTransOpList
                'dgvDirectDTOps.Rows.Add(Item.Name, Item.Type, Item.Code, Item.Accuracy, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsLevel, Item.SourceCrsCode, Item.TargetCrsLevel, Item.TargetCrsCode, Item.Reversible, Item.ApplyReverse)
                dgvDirectDTOps.Rows.Add(Item.Name, Item.Type, Item.Code, Item.Accuracy, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsLevel, Item.SourceCrsCode, Item.TargetCrsLevel, Item.TargetCrsCode, Item.Reversible, Item.ApplyReverse, Item.MethodCode, Item.MethodName)
            Next
            DisplayDirectDatumTransMethod()
        End If
    End Sub

    Private Sub DisplayInputToWgs84TransOptions()
        'Display the coordinate operations that could be used for a datum transformation from the Input Crs to WGS 84.
        If Conversion.InputCrs.Code > -1 Then
            'dgvDirectDTOps.Rows.Clear()
            dgvInputToWgs84DTOps.Rows.Clear()
            For Each Item In Conversion.InputToWgs84TransOpList
                'dgvDirectDTOps.Rows.Add(Item.Name, Item.Type, Item.Code, Item.Accuracy, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsLevel, Item.SourceCrsCode, Item.TargetCrsLevel, Item.TargetCrsCode, Item.Reversible, Item.ApplyReverse)
                'dgvInputToWgs84DTOps.Rows.Add(Item.Name, Item.Type, Item.Code, Item.Accuracy, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsLevel, Item.SourceCrsCode, Item.TargetCrsLevel, Item.TargetCrsCode, Item.Reversible, Item.ApplyReverse)
                dgvInputToWgs84DTOps.Rows.Add(Item.Name, Item.Type, Item.Code, Item.Accuracy, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsLevel, Item.SourceCrsCode, Item.TargetCrsLevel, Item.TargetCrsCode, Item.Reversible, Item.ApplyReverse, Item.MethodCode, Item.MethodName)
            Next
            DisplayInputToWgs84DatumTransMethod()
        End If
    End Sub

    Private Sub DisplayWgs84ToOutputTransOptions()
        'Display the coordinate operations that could be used for a datum transformation from WGS 84 to the Output Crs.
        If Conversion.OutputCrs.Code > -1 Then
            'dgvDirectDTOps.Rows.Clear()
            dgvWgs84ToOutputDTOps.Rows.Clear()
            For Each Item In Conversion.OutputFromWgs84TransOpList
                'dgvDirectDTOps.Rows.Add(Item.Name, Item.Type, Item.Code, Item.Accuracy, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsLevel, Item.SourceCrsCode, Item.TargetCrsLevel, Item.TargetCrsCode, Item.Reversible, Item.ApplyReverse)
                'dgvWgs84ToOutputDTOps.Rows.Add(Item.Name, Item.Type, Item.Code, Item.Accuracy, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsLevel, Item.SourceCrsCode, Item.TargetCrsLevel, Item.TargetCrsCode, Item.Reversible, Item.ApplyReverse)
                dgvWgs84ToOutputDTOps.Rows.Add(Item.Name, Item.Type, Item.Code, Item.Accuracy, Item.Deprecated, Item.Version, Item.RevisionDate, Item.SourceCrsLevel, Item.SourceCrsCode, Item.TargetCrsLevel, Item.TargetCrsCode, Item.Reversible, Item.ApplyReverse, Item.MethodCode, Item.MethodName)
            Next
            DisplayWgs84ToOutputDatumTransMethod()
        End If
    End Sub

    Private Sub rbDtNotRequired_CheckedChanged(sender As Object, e As EventArgs) Handles rbDtNotRequired.CheckedChanged
        If rbDtNotRequired.Checked And rbDtNotRequired.Focused Then
            Conversion.DatumTrans.Type = clsDatumTrans.enumType.None
            txtAccuracy.Text = ""
        End If
    End Sub

    Private Sub rbDatumTransViaWgs84_CheckedChanged(sender As Object, e As EventArgs) Handles rbDatumTransViaWgs84.CheckedChanged

        If rbDatumTransViaWgs84.Checked And rbDatumTransViaWgs84.Focused Then
            TabControl4.SelectedIndex = 1
            Conversion.DatumTrans.Type = clsDatumTrans.enumType.ViaWgs84
            txtAccuracy.Text = Conversion.DatumTrans.InputToWgs84CoordOp.Accuracy + Conversion.DatumTrans.Wgs84ToOutputCoordOp.Accuracy
        End If
    End Sub

    Private Sub rbDirectDatumTrans_CheckedChanged(sender As Object, e As EventArgs) Handles rbDirectDatumTrans.CheckedChanged

        If rbDirectDatumTrans.Checked And rbDirectDatumTrans.Focused Then
            TabControl4.SelectedIndex = 0
            Conversion.DatumTrans.Type = clsDatumTrans.enumType.Direct
            txtAccuracy.Text = Conversion.DatumTrans.DirectCoordOp.Accuracy
        End If
    End Sub

    Private Sub dgvDatumTransOps_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DisplayDirectDatumTransMethod()
        'Display the Direct Datum Transformation method.

        If Conversion.DatumTrans.DirectCoordOpMethod.Code = -1 Then
            'No Direct Datum Transformation Method defined.
            txtDirectDTOpName.Text = ""
            txtDirectDTOpCode.Text = ""
            txtDirectDTOpAccuracy.Text = ""
            txtDirectDTOpRemarks.Text = ""
            txtDirectDTApplyReverse.Text = ""
            txtDirectDTMethodName.Text = ""
            txtDirectDTMethodCode.Text = ""
            txtDirectDTMethodReversible.Text = ""
            txtInputToWgs84DTFormula.Text = ""
            txtInputToWgs84DTExample.Text = ""
            txtDirectDTMethodRemarks.Text = ""
            dgvInputToWgs84DTParams.Rows.Clear()
        Else
            txtDirectDTOpName.Text = Conversion.DatumTrans.DirectCoordOp.Name
            txtDirectDTOpCode.Text = Conversion.DatumTrans.DirectCoordOp.Code
            txtDirectDTOpAccuracy.Text = Conversion.DatumTrans.DirectCoordOp.Accuracy
            txtDirectDTOpRemarks.Text = Conversion.DatumTrans.DirectCoordOp.Remarks
            txtDirectDTApplyReverse.Text = Conversion.DatumTrans.DirectMethodApplyReverse
            txtDirectDTMethodName.Text = Conversion.DatumTrans.DirectCoordOpMethod.Name
            txtDirectDTMethodCode.Text = Conversion.DatumTrans.DirectCoordOpMethod.Code
            txtDirectDTMethodReversible.Text = Conversion.DatumTrans.DirectCoordOpMethod.ReverseOp

            txtDirectDTFormula.Text = Conversion.DatumTrans.DirectCoordOpMethod.Formula
            txtDirectDTExample.Text = Conversion.DatumTrans.DirectCoordOpMethod.Example

            txtDirectDTMethodRemarks.Text = Conversion.DatumTrans.DirectCoordOpMethod.Remarks

            dgvDirectDTParams.Rows.Clear()

            Dim ParamNo As Integer
            Dim NParams As Integer = Conversion.DatumTrans.DirectCoordOpParamList.Count
            Try
                If NParams > 0 Then
                    Dim UomName As String
                    For ParamNo = 0 To NParams - 1
                        If Conversion.DatumTrans.DirectCoordOpParamValList(ParamNo).UomCode = -1 Then
                            UomName = ""
                        Else
                            UomName = Conversion.UnitOfMeas(Conversion.DatumTrans.DirectCoordOpParamValList(ParamNo).UomCode).Name
                        End If
                        dgvDirectDTParams.Rows.Add(Conversion.DatumTrans.DirectCoordOpParamList(ParamNo).Code,
                                               Conversion.DatumTrans.DirectCoordOpParamUseList(ParamNo).SortOrder,
                                               Conversion.DatumTrans.DirectCoordOpParamList(ParamNo).Name,
                                               Conversion.DatumTrans.DirectCoordOpParamUseList(ParamNo).SignReversal,
                                               Conversion.DatumTrans.DirectCoordOpParamValList(ParamNo).ParameterValue,
                                               UomName,
                                               Conversion.DatumTrans.DirectCoordOpParamList(ParamNo).Description)
                    Next
                End If
            Catch ex As Exception
                Main.Message.AddWarning(ex.Message & vbCrLf)
            End Try
            dgvDirectDTParams.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            dgvDirectDTParams.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvDirectDTParams.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvDirectDTParams.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvDirectDTParams.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvDirectDTParams.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvDirectDTParams.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

            dgvDirectDTParams.Columns(6).Width = 640

            dgvDirectDTParams.Columns(6).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvDirectDTParams.AutoResizeRows()

            dgvDirectDTParams.AllowUserToAddRows = False

        End If
    End Sub

    Private Sub DisplayInputToWgs84DatumTransMethod()
        'Display the Input To WGS 84 Datum Transformation method.

        If Conversion.DatumTrans.InputToWgs84CoordOpMethod.Code = -1 Then
            'No Input to WGS 84 Datum Transformation Method defined.
            txtInputToWgs84DTOpName.Text = ""
            txtInputToWgs84DTOpCode.Text = ""
            txtInputToWgs84DTOpAccuracy.Text = ""
            txtAccuracy.Text = ""
            txtInputToWgs84DTOpRemarks.Text = ""

            txtInputToWgs84DTApplyReverse.Text = ""
            txtInputToWgs84DTMethodName.Text = ""
            txtInputToWgs84DTMethodCode.Text = ""
            txtInputToWgs84DTMethodReversible.Text = ""
            txtInputToWgs84DTMethodRemarks.Text = ""

            txtInputToWgs84DTCharCode.Text = ""
            txtInputToWgs84DTCharName.Text = ""
            txtInputToWgs84DTFormula.Text = ""
            txtInputToWgs84DTExample.Text = ""

            dgvWgs84ToOutputDTParams.Rows.Clear()
        Else
            txtInputToWgs84DTOpName.Text = Conversion.DatumTrans.InputToWgs84CoordOp.Name
            txtInputToWgs84DTOpCode.Text = Conversion.DatumTrans.InputToWgs84CoordOp.Code
            txtInputToWgs84DTOpAccuracy.Text = Conversion.DatumTrans.InputToWgs84CoordOp.Accuracy
            txtInputToWgs84DTOpRemarks.Text = Conversion.DatumTrans.InputToWgs84CoordOp.Remarks

            If Conversion.DatumTrans.Wgs84ToOutputCoordOpMethod.Code = -1 Then
                txtAccuracy.Text = ""
            Else
                txtAccuracy.Text = Conversion.DatumTrans.InputToWgs84CoordOp.Accuracy + Conversion.DatumTrans.Wgs84ToOutputCoordOp.Accuracy
            End If

            txtInputToWgs84DTApplyReverse.Text = Conversion.DatumTrans.InputToWgs84MethodApplyReverse
            txtInputToWgs84DTMethodName.Text = Conversion.DatumTrans.InputToWgs84CoordOpMethod.Name
            txtInputToWgs84DTMethodCode.Text = Conversion.DatumTrans.InputToWgs84CoordOpMethod.Code
            txtInputToWgs84DTMethodReversible.Text = Conversion.DatumTrans.InputToWgs84CoordOpMethod.ReverseOp
            txtInputToWgs84DTMethodRemarks.Text = Conversion.DatumTrans.InputToWgs84CoordOpMethod.Remarks

            txtInputToWgs84DTCharCode.Text = ""
            txtInputToWgs84DTCharName.Text = ""
            txtInputToWgs84DTFormula.Text = Conversion.DatumTrans.InputToWgs84CoordOpMethod.Formula
            txtInputToWgs84DTExample.Text = Conversion.DatumTrans.InputToWgs84CoordOpMethod.Example

            dgvInputToWgs84DTParams.Rows.Clear()

            Dim ParamNo As Integer
            Dim NParams As Integer = Conversion.DatumTrans.InputToWgs84CoordOpParamList.Count
            Try
                If NParams > 0 Then
                    Dim UomName As String
                    For ParamNo = 0 To NParams - 1
                        If Conversion.DatumTrans.InputToWgs84CoordOpParamValList(ParamNo).UomCode = -1 Then
                            UomName = ""
                        Else
                            UomName = Conversion.UnitOfMeas(Conversion.DatumTrans.InputToWgs84CoordOpParamValList(ParamNo).UomCode).Name
                        End If
                        'dgvWgs84ToOutputDTParams.Rows.Add(Conversion.DatumTrans.InputToWgs84CoordOpParamList(ParamNo).Code,
                        '                       Conversion.DatumTrans.InputToWgs84CoordOpParamUseList(ParamNo).SortOrder,
                        '                       Conversion.DatumTrans.InputToWgs84CoordOpParamList(ParamNo).Name,
                        '                       Conversion.DatumTrans.InputToWgs84CoordOpParamUseList(ParamNo).SignReversal,
                        '                       Conversion.DatumTrans.InputToWgs84CoordOpParamValList(ParamNo).ParameterValue,
                        '                       UomName,
                        '                       Conversion.DatumTrans.InputToWgs84CoordOpParamList(ParamNo).Description)
                        dgvInputToWgs84DTParams.Rows.Add(Conversion.DatumTrans.InputToWgs84CoordOpParamList(ParamNo).Code,
                       Conversion.DatumTrans.InputToWgs84CoordOpParamUseList(ParamNo).SortOrder,
                       Conversion.DatumTrans.InputToWgs84CoordOpParamList(ParamNo).Name,
                       Conversion.DatumTrans.InputToWgs84CoordOpParamUseList(ParamNo).SignReversal,
                       Conversion.DatumTrans.InputToWgs84CoordOpParamValList(ParamNo).ParameterValue,
                       UomName,
                       Conversion.DatumTrans.InputToWgs84CoordOpParamList(ParamNo).Description)
                    Next
                End If
            Catch ex As Exception
                Main.Message.AddWarning(ex.Message & vbCrLf)
            End Try
            dgvInputToWgs84DTParams.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            dgvInputToWgs84DTParams.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvInputToWgs84DTParams.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvInputToWgs84DTParams.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvInputToWgs84DTParams.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvInputToWgs84DTParams.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvInputToWgs84DTParams.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

            dgvInputToWgs84DTParams.Columns(6).Width = 640

            dgvInputToWgs84DTParams.Columns(6).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvInputToWgs84DTParams.AutoResizeRows()

            dgvInputToWgs84DTParams.AllowUserToAddRows = False
        End If
    End Sub

    Private Sub DisplayWgs84ToOutputDatumTransMethod()
        'Display the Direct Datum Transformation method.

        If Conversion.DatumTrans.Wgs84ToOutputCoordOpMethod.Code = -1 Then
            'No Direct Datum Transformation Method defined.
            txtWgs84ToOutputDTOpName.Text = ""
            'txtWgs84ToOutputOpName.Text = ""
            txtWgs84ToOutputDTOpCode.Text = ""
            txtWgs84ToOutputDTOpAccuracy.Text = ""
            'txtWgs84ToOutputOpAccuracy.Text = ""
            txtAccuracy.Text = ""
            txtWgs84ToOutputDTOpRemarks.Text = ""

            txtWgs84ToOutputDTApplyReverse.Text = ""
            txtWgs84ToOutputDTMethodName.Text = ""
            'txtWgs84ToOutputMethodName.Text = ""
            txtWgs84ToOutputDTMethodCode.Text = ""
            txtWgs84ToOutputDTMethodReversible.Text = ""
            txtWgs84ToOutputDTMethodRemarks.Text = ""

            txtWgs84ToOutputDTCharCode.Text = ""
            txtWgs84ToOutputDTCharName.Text = ""
            txtWgs84ToOutputDTFormula.Text = ""
            txtWgs84ToOutputDTExample.Text = ""

            dgvWgs84ToOutputDTParams.Rows.Clear()
        Else
            txtWgs84ToOutputDTOpName.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOp.Name
            'txtWgs84ToOutputOpName.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOp.Name 'Display the Operation name on the Point Conversion \ Coordinate Type Conversion tab.
            txtWgs84ToOutputDTOpCode.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOp.Code
            txtWgs84ToOutputDTOpAccuracy.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOp.Accuracy
            'txtWgs84ToOutputOpAccuracy.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOp.Accuracy  'Display the Operation accuracy on the Point Conversion \ Coordinate Type Conversion tab.
            txtWgs84ToOutputDTOpRemarks.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOp.Remarks

            If Conversion.DatumTrans.InputToWgs84CoordOpMethod.Code = -1 Then
                txtAccuracy.Text = ""
            Else
                txtAccuracy.Text = Conversion.DatumTrans.InputToWgs84CoordOp.Accuracy + Conversion.DatumTrans.Wgs84ToOutputCoordOp.Accuracy
            End If

            txtWgs84ToOutputDTApplyReverse.Text = Conversion.DatumTrans.Wgs84ToOutputMethodApplyReverse
            txtWgs84ToOutputDTMethodName.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOpMethod.Name
            'txtWgs84ToOutputMethodName.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOpMethod.Name 'Display the Method name on the Point Conversion \ Coordinate Type Conversion tab.
            txtWgs84ToOutputDTMethodCode.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOpMethod.Code
            txtWgs84ToOutputDTMethodReversible.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOpMethod.ReverseOp
            txtWgs84ToOutputDTMethodRemarks.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOpMethod.Remarks

            txtWgs84ToOutputDTFormula.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOpMethod.Formula
            txtWgs84ToOutputDTExample.Text = Conversion.DatumTrans.Wgs84ToOutputCoordOpMethod.Example

            dgvWgs84ToOutputDTParams.Rows.Clear()

            Dim ParamNo As Integer
            Dim NParams As Integer = Conversion.DatumTrans.Wgs84ToOutputCoordOpParamList.Count
            Try
                If NParams > 0 Then
                    Dim UomName As String
                    For ParamNo = 0 To NParams - 1
                        If Conversion.DatumTrans.Wgs84ToOutputCoordOpParamValList(ParamNo).UomCode = -1 Then
                            UomName = ""
                        Else
                            UomName = Conversion.UnitOfMeas(Conversion.DatumTrans.Wgs84ToOutputCoordOpParamValList(ParamNo).UomCode).Name
                        End If
                        dgvWgs84ToOutputDTParams.Rows.Add(Conversion.DatumTrans.Wgs84ToOutputCoordOpParamList(ParamNo).Code,
                                               Conversion.DatumTrans.Wgs84ToOutputCoordOpParamUseList(ParamNo).SortOrder,
                                               Conversion.DatumTrans.Wgs84ToOutputCoordOpParamList(ParamNo).Name,
                                               Conversion.DatumTrans.Wgs84ToOutputCoordOpParamUseList(ParamNo).SignReversal,
                                               Conversion.DatumTrans.Wgs84ToOutputCoordOpParamValList(ParamNo).ParameterValue,
                                               UomName,
                                               Conversion.DatumTrans.Wgs84ToOutputCoordOpParamList(ParamNo).Description)
                    Next
                End If
            Catch ex As Exception
                Main.Message.AddWarning(ex.Message & vbCrLf)
            End Try
            dgvWgs84ToOutputDTParams.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            dgvWgs84ToOutputDTParams.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvWgs84ToOutputDTParams.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvWgs84ToOutputDTParams.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvWgs84ToOutputDTParams.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvWgs84ToOutputDTParams.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvWgs84ToOutputDTParams.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

            dgvWgs84ToOutputDTParams.Columns(6).Width = 640

            dgvWgs84ToOutputDTParams.Columns(6).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvWgs84ToOutputDTParams.AutoResizeRows()

            dgvWgs84ToOutputDTParams.AllowUserToAddRows = False

        End If
    End Sub

    Private Sub dgvDirectDTOps_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDirectDTOps.CellContentClick

        If dgvDirectDTOps.SelectedRows.Count = 0 Then
            Main.Message.AddWarning("Please select a Datum Transformation operation." & vbCrLf)
        ElseIf dgvDirectDTOps.SelectedRows.Count = 1 Then
            Dim SelRow As Integer = dgvDirectDTOps.SelectedRows(0).Index
            Dim OpCode As Integer
            OpCode = dgvDirectDTOps.Rows(SelRow).Cells(2).Value
            Dim ApplyReverse As Boolean
            ApplyReverse = dgvDirectDTOps.Rows(SelRow).Cells(12).Value

            Conversion.DatumTrans.GetDirectDatumTransCoordOp(OpCode)
            Conversion.DatumTrans.DirectMethodApplyReverse = ApplyReverse
            DisplayDirectDatumTransMethod()
        Else
            Main.Message.AddWarning("Please select a single Datum Transformation operation." & vbCrLf)
        End If
    End Sub

    Private Sub dgvInputToWgs84DTOps_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvInputToWgs84DTOps.CellContentClick

        If dgvInputToWgs84DTOps.SelectedRows.Count = 0 Then
            Main.Message.AddWarning("Please select a Datum Transformation operation." & vbCrLf)
        ElseIf dgvInputToWgs84DTOps.SelectedRows.Count = 1 Then
            Dim SelRow As Integer = dgvInputToWgs84DTOps.SelectedRows(0).Index
            Dim OpCode As Integer
            OpCode = dgvInputToWgs84DTOps.Rows(SelRow).Cells(2).Value
            Dim ApplyReverse As Boolean
            ApplyReverse = dgvInputToWgs84DTOps.Rows(SelRow).Cells(12).Value

            Conversion.DatumTrans.GetInputToWgs84DatumTransCoordOp(OpCode)
            Conversion.DatumTrans.InputToWgs84MethodApplyReverse = ApplyReverse
            DisplayInputToWgs84DatumTransMethod()
        Else
            Main.Message.AddWarning("Please select a single Datum Transformation operation." & vbCrLf)
        End If
    End Sub


    Private Sub dgvWgs84ToOutputDTOps_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvWgs84ToOutputDTOps.CellContentClick

        If dgvWgs84ToOutputDTOps.SelectedRows.Count = 0 Then
            Main.Message.AddWarning("Please select a Datum Transformation operation." & vbCrLf)
        ElseIf dgvWgs84ToOutputDTOps.SelectedRows.Count = 1 Then
            Dim SelRow As Integer = dgvWgs84ToOutputDTOps.SelectedRows(0).Index
            Dim OpCode As Integer
            OpCode = dgvWgs84ToOutputDTOps.Rows(SelRow).Cells(2).Value
            Dim ApplyReverse As Boolean
            ApplyReverse = dgvWgs84ToOutputDTOps.Rows(SelRow).Cells(12).Value

            Conversion.DatumTrans.GetWgs84ToOutputDatumTransCoordOp(OpCode)
            Conversion.DatumTrans.Wgs84ToOutputMethodApplyReverse = ApplyReverse
            DisplayWgs84ToOutputDatumTransMethod()
        Else
            Main.Message.AddWarning("Please select a single Datum Transformation operation." & vbCrLf)
        End If
    End Sub

    Private Sub TabControl4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl4.SelectedIndexChanged

        If TabControl4.SelectedIndex = 0 Then 'Direct Datum Transformation
            rbDirectDatumTrans.Checked = True
            SplitContainer12.SplitterDistance = SplitDist12
            SplitContainer13.SplitterDistance = SplitDist13
            SplitContainer14.SplitterDistance = SplitDist14
        Else
            rbDatumTransViaWgs84.Checked = True
        End If
    End Sub

    Private Sub TabControl5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl5.SelectedIndexChanged

        If TabControl5.SelectedIndex = 0 Then 'Input ro WGS 84
            SplitContainer7.SplitterDistance = SplitDist7
            SplitContainer10.SplitterDistance = SplitDist10
            SplitContainer11.SplitterDistance = SplitDist11
        Else 'WGS 84 to Putput
            SplitContainer15.SplitterDistance = SplitDist15
            SplitContainer8.SplitterDistance = SplitDist8
            SplitContainer9.SplitterDistance = SplitDist9
        End If
    End Sub

    Private Sub btnFormatHelp2_Click(sender As Object, e As EventArgs) Handles btnFormatHelp2.Click
        'Show Format information.
        MessageBox.Show("Format string examples:" & vbCrLf & "N4 - Number displayed with thousands separator and 4 decimal places" & vbCrLf & "F4 - Number displayed with 4 decimal places.", "Number Formatting")
    End Sub

    Private Sub SplitContainer1_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer1.SplitterMoved
        SplitDist1 = SplitContainer1.SplitterDistance
    End Sub

    Private Sub SplitContainer2_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer2.SplitterMoved
        SplitDist2 = SplitContainer2.SplitterDistance
    End Sub

    Private Sub SplitContainer3_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer3.SplitterMoved
        SplitDist3 = SplitContainer3.SplitterDistance
    End Sub

    Private Sub SplitContainer4_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer4.SplitterMoved
        SplitDist4 = SplitContainer4.SplitterDistance
    End Sub

    Private Sub SplitContainer5_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer5.SplitterMoved
        SplitDist5 = SplitContainer5.SplitterDistance
    End Sub

    Private Sub SplitContainer6_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer6.SplitterMoved
        SplitDist6 = SplitContainer6.SplitterDistance
    End Sub

    Private Sub SplitContainer7_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer7.SplitterMoved
        SplitDist7 = SplitContainer7.SplitterDistance
    End Sub

    Private Sub SplitContainer8_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer8.SplitterMoved
        SplitDist8 = SplitContainer8.SplitterDistance
    End Sub

    Private Sub SplitContainer9_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer9.SplitterMoved
        SplitDist9 = SplitContainer9.SplitterDistance
    End Sub

    Private Sub SplitContainer10_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer10.SplitterMoved
        SplitDist10 = SplitContainer10.SplitterDistance
    End Sub

    Private Sub SplitContainer11_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer11.SplitterMoved
        SplitDist11 = SplitContainer11.SplitterDistance
    End Sub

    Private Sub SplitContainer12_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer12.SplitterMoved
        SplitDist12 = SplitContainer12.SplitterDistance
    End Sub

    Private Sub SplitContainer13_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer13.SplitterMoved
        SplitDist13 = SplitContainer13.SplitterDistance
    End Sub

    Private Sub SplitContainer14_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer14.SplitterMoved
        SplitDist14 = SplitContainer14.SplitterDistance
    End Sub

    Private Sub SplitContainer15_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer15.SplitterMoved
        SplitDist15 = SplitContainer15.SplitterDistance
    End Sub


    Private Sub rbEnterInputEastNorth_CheckedChanged(sender As Object, e As EventArgs) Handles rbEnterInputEastNorth.CheckedChanged
        If rbEnterInputEastNorth.Checked Then UpdateDatumTransTableInput()
    End Sub

    Private Sub rbEnterInputLongLat_CheckedChanged(sender As Object, e As EventArgs) Handles rbEnterInputLongLat.CheckedChanged
        If rbEnterInputLongLat.Checked Then UpdateDatumTransTableInput()
    End Sub

    Private Sub rbEnterInputXYZ_CheckedChanged(sender As Object, e As EventArgs) Handles rbEnterInputXYZ.CheckedChanged
        If rbEnterInputXYZ.Checked Then UpdateDatumTransTableInput()
    End Sub



    Private Sub chkShowPointNumber_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowPointNumber.CheckedChanged
        'UpdateDatumTransTable()
        If chkShowPointNumber.Checked Then 'Show the Point Number column.
            dgvConversion.Columns(0).Visible = True
        Else 'Remove the Point Number column
            dgvConversion.Columns(0).Visible = False
        End If
    End Sub

    Private Sub chkShowPointName_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowPointName.CheckedChanged
        'UpdateDatumTransTable()
        If chkShowPointName.Checked Then 'Show the Point Name column.
            dgvConversion.Columns(1).Visible = True
        Else 'Hide the Point Name column
            dgvConversion.Columns(1).Visible = False
        End If
    End Sub

    Private Sub chkShowPointDescription_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowPointDescription.CheckedChanged
        'UpdateDatumTransTable()
        If chkShowPointDescription.Checked Then 'Show the Point Description column.
            dgvConversion.Columns(2).Visible = True
        Else 'Hide the Point Description column
            dgvConversion.Columns(2).Visible = False
        End If
    End Sub

    Private Sub chkShowInputEastNorth_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowInputEastNorth.CheckedChanged
        'UpdateDatumTransTable()
        If chkShowInputEastNorth.Checked Then 'Show the Input Easting and Northing columns.
            dgvConversion.Columns(3).Visible = True
            dgvConversion.Columns(4).Visible = True
        Else 'Hide the Input Easting and Northing columns.
            dgvConversion.Columns(3).Visible = False
            dgvConversion.Columns(4).Visible = False
        End If
    End Sub

    Private Sub chkShowInputLongLat_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowInputLongLat.CheckedChanged
        'UpdateDatumTransTable()
        If chkShowInputLongLat.Checked Then 'Show the Input Longitude, Latitude and Ellipsoidal Height columns.
            dgvConversion.Columns(5).Visible = True
            dgvConversion.Columns(6).Visible = True
            dgvConversion.Columns(7).Visible = True
        Else 'Hide the Input Longitude, Latitude and Ellipsoidal Height columns.
            dgvConversion.Columns(5).Visible = False
            dgvConversion.Columns(6).Visible = False
            dgvConversion.Columns(7).Visible = False
        End If
    End Sub

    Private Sub chkShowInputXYZ_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowInputXYZ.CheckedChanged
        'UpdateDatumTransTable()
        If chkShowInputXYZ.Checked Then 'Show the Input X, Y and Z columns.
            dgvConversion.Columns(8).Visible = True
            dgvConversion.Columns(9).Visible = True
            dgvConversion.Columns(10).Visible = True
        Else 'Hide the Input X, Y and Z columns.
            dgvConversion.Columns(8).Visible = False
            dgvConversion.Columns(9).Visible = False
            dgvConversion.Columns(10).Visible = False
        End If
    End Sub

    Private Sub chkShowWgs84XYZ_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowWgs84XYZ.CheckedChanged
        'UpdateDatumTransTable()
        If chkShowWgs84XYZ.Checked Then 'Show the WGS 84 X, Y and Z columns.
            dgvConversion.Columns(11).Visible = True
            dgvConversion.Columns(12).Visible = True
            dgvConversion.Columns(13).Visible = True
        Else 'Hide the WGS 84 X, Y and Z columns.
            dgvConversion.Columns(11).Visible = False
            dgvConversion.Columns(12).Visible = False
            dgvConversion.Columns(13).Visible = False
        End If
    End Sub

    Private Sub chkShowOutputXYZ_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowOutputXYZ.CheckedChanged
        'UpdateDatumTransTable()
        If chkShowOutputXYZ.Checked Then 'Show the Output X, Y and Z columns.
            dgvConversion.Columns(14).Visible = True
            dgvConversion.Columns(15).Visible = True
            dgvConversion.Columns(16).Visible = True
        Else 'Hide the Output X, Y and Z columns.
            dgvConversion.Columns(14).Visible = False
            dgvConversion.Columns(15).Visible = False
            dgvConversion.Columns(16).Visible = False
        End If
    End Sub

    Private Sub chkShowOutputLongLat_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowOutputLongLat.CheckedChanged
        'UpdateDatumTransTable()
        If chkShowOutputLongLat.Checked Then 'Show the Output Longitude, Latitude and Ellipsoidal Height columns.
            dgvConversion.Columns(17).Visible = True
            dgvConversion.Columns(18).Visible = True
            dgvConversion.Columns(19).Visible = True
        Else 'Hide the Output Longitude, Latitude and Ellipsoidal Height columns.
            dgvConversion.Columns(17).Visible = False
            dgvConversion.Columns(18).Visible = False
            dgvConversion.Columns(19).Visible = False
        End If
    End Sub

    Private Sub chkShowOutputEastNorth_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowOutputEastNorth.CheckedChanged
        'UpdateDatumTransTable()
        If chkShowOutputEastNorth.Checked Then 'Show the Output Easting and Northing columns.
            dgvConversion.Columns(20).Visible = True
            dgvConversion.Columns(21).Visible = True
        Else 'Hide the Output Easting and Northing columns.
            dgvConversion.Columns(20).Visible = False
            dgvConversion.Columns(21).Visible = False
        End If
    End Sub



    Private Sub dgvConversion_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvConversion.CellContentClick



    End Sub

    Private Sub dgvConversion_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvConversion.CellClick

        If e.RowIndex < 0 Then
            'Header row selected - Ignore.
        Else
            Dim ColHeader As String = dgvConversion.Columns(e.ColumnIndex).HeaderText

            Select Case ColHeader
                Case "Point Number"

                Case "Point Name"

                Case "Point Description"

                Case "Input Easting"

                Case "Input Northing"

                Case "Input Longitude"
                    Conversion.Angle.SetAngle(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                    Conversion.Angle.DecDegreeDecPlaces = txtDegreeDecPlaces.Text
                    Conversion.Angle.SecondsDecPlaces = txtSecondsDecPlaces.Text
                    'txtDecDegree.Text = Conversion.Angle.DecimalDegrees
                    txtDecDegree.Text = Conversion.Angle.FormattedDecimalDegrees
                    If Conversion.Angle.Sign = clsAngle.AngleSign.Negative Then txtAngleSign.Text = "-" Else txtAngleSign.Text = "+"
                    txtAngleDegrees.Text = Conversion.Angle.Degrees
                    txtAngleMinutes.Text = Conversion.Angle.Minutes
                    'txtAngleSeconds.Text = Conversion.Angle.Seconds
                    txtAngleSeconds.Text = Conversion.Angle.FormattedSeconds
                Case "Input Latitude"
                    Conversion.Angle.SetAngle(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                    Conversion.Angle.DecDegreeDecPlaces = txtDegreeDecPlaces.Text
                    Conversion.Angle.SecondsDecPlaces = txtSecondsDecPlaces.Text
                    txtDecDegree.Text = Conversion.Angle.FormattedDecimalDegrees
                    If Conversion.Angle.Sign = clsAngle.AngleSign.Negative Then txtAngleSign.Text = "-" Else txtAngleSign.Text = "+"
                    txtAngleDegrees.Text = Conversion.Angle.Degrees
                    txtAngleMinutes.Text = Conversion.Angle.Minutes
                    txtAngleSeconds.Text = Conversion.Angle.FormattedSeconds
                Case "Input Ellipsoidal Height"

                Case "Input X"

                Case "Input Y"

                Case "Input Z"

                Case "WGS 84 X"

                Case "WGS 84 Y"

                Case "WGS 84 Z"

                Case "Output X"

                Case "Output Y"

                Case "Output Z"

                Case "Output Longitude"
                    Conversion.Angle.SetAngle(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                    'txtDecDegree.Text = Conversion.Angle.DecimalDegrees
                    txtDecDegree.Text = Conversion.Angle.FormattedDecimalDegrees
                    If Conversion.Angle.Sign = clsAngle.AngleSign.Negative Then txtAngleSign.Text = "-" Else txtAngleSign.Text = "+"
                    txtAngleDegrees.Text = Conversion.Angle.Degrees
                    txtAngleMinutes.Text = Conversion.Angle.Minutes
                    'txtAngleSeconds.Text = Conversion.Angle.Seconds
                    txtAngleSeconds.Text = Conversion.Angle.FormattedSeconds
                Case "Output Latitude"
                    Conversion.Angle.SetAngle(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                    'txtDecDegree.Text = Conversion.Angle.DecimalDegrees
                    txtDecDegree.Text = Conversion.Angle.FormattedDecimalDegrees
                    If Conversion.Angle.Sign = clsAngle.AngleSign.Negative Then txtAngleSign.Text = "-" Else txtAngleSign.Text = "+"
                    txtAngleDegrees.Text = Conversion.Angle.Degrees
                    txtAngleMinutes.Text = Conversion.Angle.Minutes
                    'txtAngleSeconds.Text = Conversion.Angle.Seconds
                    txtAngleSeconds.Text = Conversion.Angle.FormattedSeconds
                Case "Output Ellipsoidal Height"

                Case "Output Easting"

                Case "Output Northing"

                Case Else

            End Select
        End If
    End Sub

    Private Sub dgvConversion_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvConversion.CellEndEdit
        'The cell has been edited - update the coordinates.
        Dim ColHeader As String = dgvConversion.Columns(e.ColumnIndex).HeaderText
        Select Case ColHeader
            Case "Point Number"
                'The Point number has been edited

            Case "Point Name"
                'The Point name has been edited

            Case "Point Description"
                'The Point description has been edited

            Case "Input Easting"
                If rbEnterInputEastNorth.Checked Then 'This is a valid entry coordinate
                    'Check that the Northing value is valid:
                    If dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).ToString.Trim = "" Then  'There is no Northing value
                        'New coordinate values can not be calculated.
                    Else
                        Conversion.InputCrs.Coord.SetNorthing(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value, Coordinate.UpdateMode.None) 'Set the Input Northing value
                        Conversion.InputCrs.Coord.SetEasting(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Easting value and Update all the Input and Output coordinate types
                        UpdateDatumTransTable(e.RowIndex)
                    End If
                End If
            Case "Input Northing"
                If rbEnterInputEastNorth.Checked Then 'This is a valid entry coordinate
                    'Check that the Easting value is valid:
                    If dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).ToString.Trim = "" Then  'There is no Easting value
                        'New coordinate values can not be calculated.
                    Else
                        Conversion.InputCrs.Coord.SetEasting(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).Value, Coordinate.UpdateMode.None) 'Set the Input Easting value
                        Conversion.InputCrs.Coord.SetNorthing(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Northing value and Update all the Input and Output coordinate types
                        UpdateDatumTransTable(e.RowIndex)
                    End If
                End If
            Case "Input Longitude"
                If rbEnterInputLongLat.Checked Then
                    'Check if the Ellipsoidal Height has been defined:
                    If dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 2).ToString.Trim = "" Then dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 2).Value = 0
                    'Check that the Latitude value is valid:
                    If dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).ToString.Trim = "" Then  'There is no Latitude value
                        'New coordinate values can not be calculated.
                    Else
                        Conversion.InputCrs.Coord.SetLatitude(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value, Coordinate.UpdateMode.None) 'Set the Input Latitude value
                        Conversion.InputCrs.Coord.SetEllipsoidalHeight(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 2).Value, Coordinate.UpdateMode.None) 'Set the Input Ellipsoidal Height value
                        Conversion.InputCrs.Coord.SetLongitude(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Longitude value and Update all the Input and Output coordinate types
                        UpdateDatumTransTable(e.RowIndex)
                    End If
                End If
            Case "Input Latitude"
                If rbEnterInputLongLat.Checked Then
                    'Check if the Ellipsoidal Height has been defined:
                    If dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).ToString.Trim = "" Then dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value = 0
                    'Check that the Longitude value is valid:
                    If dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).ToString.Trim = "" Then  'There is no Longitude value
                        'New coordinate values can not be calculated.
                    Else
                        Conversion.InputCrs.Coord.SetLongitude(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).Value, Coordinate.UpdateMode.None) 'Set the Input Longitude value
                        Conversion.InputCrs.Coord.SetEllipsoidalHeight(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value, Coordinate.UpdateMode.None) 'Set the Input Ellipsoidal Height value
                        Conversion.InputCrs.Coord.SetLatitude(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Latitude value and Update all the Input and Output coordinate types
                        UpdateDatumTransTable(e.RowIndex)
                    End If
                End If
            Case "Input Ellipsoidal Height"
                If rbEnterInputLongLat.Checked Then
                    'Check that the Longitude and Latitude values are valid:
                    If dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).ToString.Trim = "" Or dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 2).ToString.Trim = "" Then  'There is no Longitude or Latitude value
                        'New coordinate values can not be calculated.
                    Else
                        Conversion.InputCrs.Coord.SetLongitude(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 2).Value, Coordinate.UpdateMode.None)
                        Conversion.InputCrs.Coord.SetLatitude(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).Value, Coordinate.UpdateMode.None)
                        Conversion.InputCrs.Coord.SetEllipsoidalHeight(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.InputOutputAll)
                        If Conversion.InputCrs.Kind = CoordRefSystem.CrsKind.projected Then Conversion.InputCrs.LongLatToEastNorth()
                        Conversion.InputCrs.LongLatEllHtToXYZ()
                        Conversion.DatumTrans.InputToOutput()
                        Conversion.OutputCrs.XYZToLongLatEllHt()
                        If Conversion.OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then Conversion.OutputCrs.LongLatToEastNorth()
                        UpdateDatumTransTable(e.RowIndex)
                    End If
                End If
            Case "Input X"
                If rbEnterInputXYZ.Checked Then
                    'Check if the Y and Z values are valid:
                    If dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).ToString.Trim = "" Or dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 2).ToString.Trim = "" Then  'There is no Y or Z value
                        'New coordinate values can not be calculated.
                    Else
                        Conversion.InputCrs.Coord.SetY(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value, Coordinate.UpdateMode.None) 'Set the Input Y value
                        Conversion.InputCrs.Coord.SetZ(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 2).Value, Coordinate.UpdateMode.None) 'Set the Input Z value
                        Conversion.InputCrs.Coord.SetX(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input X value and Update all the Input and Output coordinate types
                        UpdateDatumTransTable(e.RowIndex)
                    End If
                End If
            Case "Input Y"
                If rbEnterInputXYZ.Checked Then
                    'Check if the X and Z values are valid:
                    If dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).ToString.Trim = "" Or dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).ToString.Trim = "" Then  'There is no X or Z value
                        'New coordinate values can not be calculated.
                    Else
                        Conversion.InputCrs.Coord.SetX(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).Value, Coordinate.UpdateMode.None) 'Set the Input X value
                        Conversion.InputCrs.Coord.SetZ(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value, Coordinate.UpdateMode.None) 'Set the Input Z value
                        Conversion.InputCrs.Coord.SetY(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Y value and Update all the Input and Output coordinate types
                        UpdateDatumTransTable(e.RowIndex)
                    End If
                End If
            Case "Input Z"
                If rbEnterInputXYZ.Checked Then
                    'Check if the X and Y values are valid:
                    If dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 2).ToString.Trim = "" Or dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).ToString.Trim = "" Then  'There is no X or Y value
                        'New coordinate values can not be calculated.
                    Else
                        Conversion.InputCrs.Coord.SetX(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 2).Value, Coordinate.UpdateMode.None) 'Update the Input X value
                        Conversion.InputCrs.Coord.SetY(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).Value, Coordinate.UpdateMode.None) 'Update the Input Y value
                        Conversion.InputCrs.Coord.SetZ(dgvConversion.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Z value and Update all the Input and Output coordinate types
                        UpdateDatumTransTable(e.RowIndex)
                    End If
                End If
            Case "WGS 84 X"
                'This column is not editable.
            Case "WGS 84 Y"
                 'This column is not editable.
            Case "WGS 84 Z"
                 'This column is not editable.
            Case "Output X"
                 'This column is not editable.
            Case "Output Y"
                 'This column is not editable.
            Case "Output Z"
                 'This column is not editable.
            Case "Output Longitude"
                 'This column is not editable.
            Case "Output Latitude"
                 'This column is not editable.
            Case "Output Ellipsoidal Height"
                 'This column is not editable.
            Case "Output Easting"
                 'This column is not editable.
            Case "Output Northing"
                'This column is not editable.
            Case Else

        End Select
    End Sub



    Private Sub txtProjFormat_TextChanged(sender As Object, e As EventArgs) Handles txtProjFormat.TextChanged

    End Sub

    Private Sub txtProjFormat_LostFocus(sender As Object, e As EventArgs) Handles txtProjFormat.LostFocus
        'The Projected Easting and Northing format has changed.

        Dim Format As String = txtProjFormat.Text.Trim
        'For Each Col As DataGridViewColumn In dgvConversion.Columns
        '    If Col.HeaderText = "Input Easting" Then Col.DefaultCellStyle.Format = Format
        '    If Col.Visible And Col.HeaderText = "Input Easting" Then Col.DefaultCellStyle.Format = Format
        '    If Col.Visible And Col.HeaderText = "Input Northing" Then Col.DefaultCellStyle.Format = Format
        '    If Col.Visible And Col.Visible And Col.HeaderText = "Output Easting" Then Col.DefaultCellStyle.Format = Format
        '    If Col.Visible And Col.HeaderText = "Output Northing" Then Col.DefaultCellStyle.Format = Format
        'Next
        dgvConversion.Columns(3).DefaultCellStyle.Format = Format
        dgvConversion.Columns(4).DefaultCellStyle.Format = Format
        dgvConversion.Columns(20).DefaultCellStyle.Format = Format
        dgvConversion.Columns(21).DefaultCellStyle.Format = Format

    End Sub

    Private Sub txtDegreeDecPlaces_TextChanged(sender As Object, e As EventArgs) Handles txtDegreeDecPlaces.TextChanged

    End Sub

    Private Sub txtDegreeDecPlaces_LostFocus(sender As Object, e As EventArgs) Handles txtDegreeDecPlaces.LostFocus
        'The number of Decimal Degrees decimal places has changed.
        If rbDecDegrees.Checked Then
            Try
                txtDegreeDecPlaces.Text = Int(txtDegreeDecPlaces.Text.Trim)
                Dim Format As String = "F" & txtDegreeDecPlaces.Text.Trim
                dgvConversion.Columns(5).DefaultCellStyle.Format = Format
                dgvConversion.Columns(6).DefaultCellStyle.Format = Format
                dgvConversion.Columns(17).DefaultCellStyle.Format = Format
                dgvConversion.Columns(18).DefaultCellStyle.Format = Format
            Catch ex As Exception

            End Try
        End If
    End Sub

    'Private Sub SetDegreeDecPlaces(DecPlaces As Integer)
    '    Try
    '        'txtDegreeDecPlaces.Text = Int(txtDegreeDecPlaces.Text.Trim)
    '        'Dim Format As String = "F" & txtDegreeDecPlaces.Text.Trim
    '        Dim Format As String = "F" & DecPlaces
    '        For Each Col As DataGridViewColumn In dgvConversion.Columns
    '            'If Col.HeaderText = "Input Longitude" Then Col.DefaultCellStyle.Format = Format
    '            If Col.Visible And Col.HeaderText = "Input Longitude" Then Col.DefaultCellStyle.Format = Format
    '            If Col.Visible And Col.HeaderText = "Input Latitude" Then Col.DefaultCellStyle.Format = Format
    '            If Col.Visible And Col.HeaderText = "Output Longitude" Then Col.DefaultCellStyle.Format = Format
    '            If Col.Visible And Col.HeaderText = "Output Latitude" Then Col.DefaultCellStyle.Format = Format
    '        Next
    '    Catch ex As Exception

    '    End Try
    'End Sub

    Private Sub txtCartFormat_TextChanged(sender As Object, e As EventArgs) Handles txtCartFormat.TextChanged

    End Sub

    Private Sub txtCartFormat_LostFocus(sender As Object, e As EventArgs) Handles txtCartFormat.LostFocus
        'The Cartesian coordinate format has changed.
        Dim Format As String = txtCartFormat.Text.Trim
        'For Each Col As DataGridViewColumn In dgvConversion.Columns
        '    If Col.HeaderText = "Input X" Then Col.DefaultCellStyle.Format = Format
        '    If Col.HeaderText = "Input Y" Then Col.DefaultCellStyle.Format = Format
        '    If Col.HeaderText = "Input Z" Then Col.DefaultCellStyle.Format = Format
        '    If Col.HeaderText = "Output X" Then Col.DefaultCellStyle.Format = Format
        '    If Col.HeaderText = "Output Y" Then Col.DefaultCellStyle.Format = Format
        '    If Col.HeaderText = "Output Z" Then Col.DefaultCellStyle.Format = Format
        'Next
        dgvConversion.Columns(8).DefaultCellStyle.Format = Format
        dgvConversion.Columns(9).DefaultCellStyle.Format = Format
        dgvConversion.Columns(10).DefaultCellStyle.Format = Format
        dgvConversion.Columns(11).DefaultCellStyle.Format = Format
        dgvConversion.Columns(12).DefaultCellStyle.Format = Format
        dgvConversion.Columns(13).DefaultCellStyle.Format = Format
        dgvConversion.Columns(14).DefaultCellStyle.Format = Format
        dgvConversion.Columns(15).DefaultCellStyle.Format = Format
        dgvConversion.Columns(16).DefaultCellStyle.Format = Format
    End Sub

    Private Sub txtHeightFormat_TextChanged(sender As Object, e As EventArgs) Handles txtHeightFormat.TextChanged

    End Sub

    Private Sub txtHeightFormat_LostFocus(sender As Object, e As EventArgs) Handles txtHeightFormat.LostFocus
        'The Ellipsoidal Height format has changed.
        Dim Format As String = txtHeightFormat.Text.Trim
        'For Each Col As DataGridViewColumn In dgvConversion.Columns
        '    'If Col.HeaderText = "Input Ellipsoidal Height" Then Col.DefaultCellStyle.Format = Format
        '    If Col.Visible And Col.HeaderText = "Input Ellipsoidal Height" Then Col.DefaultCellStyle.Format = Format
        '    If Col.Visible And Col.HeaderText = "Output Ellipsoidal Height" Then Col.DefaultCellStyle.Format = Format
        'Next
        dgvConversion.Columns(7).DefaultCellStyle.Format = Format
        dgvConversion.Columns(19).DefaultCellStyle.Format = Format
    End Sub

    Private Sub ApplyDatumTransFormats()
        'Apply the Dataum Transformation number formats to dgvConversion

        Dim ProjFormat As String = txtProjFormat.Text.Trim
        Dim CartFormat As String = txtCartFormat.Text.Trim
        Dim HeightFormat As String = txtHeightFormat.Text.Trim

        'For Each Col As DataGridViewColumn In dgvConversion.Columns
        '    'If Col.HeaderText = "Input Easting" Then Col.DefaultCellStyle.Format = ProjFormat
        '    If Col.Visible And Col.HeaderText = "Input Easting" Then Col.DefaultCellStyle.Format = ProjFormat
        '    If Col.Visible And Col.HeaderText = "Input Northing" Then Col.DefaultCellStyle.Format = ProjFormat
        '    If Col.Visible And Col.HeaderText = "Output Easting" Then Col.DefaultCellStyle.Format = ProjFormat
        '    If Col.Visible And Col.HeaderText = "Output Northing" Then Col.DefaultCellStyle.Format = ProjFormat

        '    If Col.Visible And Col.HeaderText = "Input Ellipsoidal Height" Then Col.DefaultCellStyle.Format = HeightFormat
        '    If Col.Visible And Col.HeaderText = "Output Ellipsoidal Height" Then Col.DefaultCellStyle.Format = HeightFormat

        '    If Col.Visible And Col.HeaderText = "Input X" Then Col.DefaultCellStyle.Format = CartFormat
        '    If Col.Visible And Col.HeaderText = "Input Y" Then Col.DefaultCellStyle.Format = CartFormat
        '    If Col.Visible And Col.HeaderText = "Input Z" Then Col.DefaultCellStyle.Format = CartFormat
        '    If Col.Visible And Col.HeaderText = "Output X" Then Col.DefaultCellStyle.Format = CartFormat
        '    If Col.Visible And Col.HeaderText = "Output Y" Then Col.DefaultCellStyle.Format = CartFormat
        '    If Col.Visible And Col.HeaderText = "Output Z" Then Col.DefaultCellStyle.Format = CartFormat
        'Next

        dgvConversion.Columns(3).DefaultCellStyle.Format = ProjFormat
        dgvConversion.Columns(4).DefaultCellStyle.Format = ProjFormat
        dgvConversion.Columns(20).DefaultCellStyle.Format = ProjFormat
        dgvConversion.Columns(21).DefaultCellStyle.Format = ProjFormat

        dgvConversion.Columns(7).DefaultCellStyle.Format = HeightFormat
        dgvConversion.Columns(19).DefaultCellStyle.Format = HeightFormat

        dgvConversion.Columns(8).DefaultCellStyle.Format = CartFormat
        dgvConversion.Columns(9).DefaultCellStyle.Format = CartFormat
        dgvConversion.Columns(10).DefaultCellStyle.Format = CartFormat
        dgvConversion.Columns(11).DefaultCellStyle.Format = CartFormat
        dgvConversion.Columns(12).DefaultCellStyle.Format = CartFormat
        dgvConversion.Columns(13).DefaultCellStyle.Format = CartFormat
        dgvConversion.Columns(14).DefaultCellStyle.Format = CartFormat
        dgvConversion.Columns(15).DefaultCellStyle.Format = CartFormat
        dgvConversion.Columns(16).DefaultCellStyle.Format = CartFormat

        If rbDecDegrees.Checked Then
            Try
                txtDegreeDecPlaces.Text = Int(txtDegreeDecPlaces.Text.Trim)
                Dim Format As String = "F" & txtDegreeDecPlaces.Text.Trim
                'For Each Col As DataGridViewColumn In dgvConversion.Columns
                '    'If Col.HeaderText = "Input Longitude" Then Col.DefaultCellStyle.Format = Format
                '    If Col.Visible And Col.HeaderText = "Input Longitude" Then Col.DefaultCellStyle.Format = Format
                '    If Col.Visible And Col.HeaderText = "Input Latitude" Then Col.DefaultCellStyle.Format = Format
                '    If Col.Visible And Col.HeaderText = "Output Longitude" Then Col.DefaultCellStyle.Format = Format
                '    If Col.Visible And Col.HeaderText = "Output Latitude" Then Col.DefaultCellStyle.Format = Format
                'Next
                dgvConversion.Columns(5).DefaultCellStyle.Format = Format
                dgvConversion.Columns(6).DefaultCellStyle.Format = Format
                dgvConversion.Columns(17).DefaultCellStyle.Format = Format
                dgvConversion.Columns(18).DefaultCellStyle.Format = Format
            Catch ex As Exception

            End Try
        Else

        End If

        dgvConversion.AutoResizeColumns()

    End Sub

    Private Sub txtSecondsDecPlaces_TextChanged(sender As Object, e As EventArgs) Handles txtSecondsDecPlaces.TextChanged

    End Sub

    Private Sub txtSecondsDecPlaces_LostFocus(sender As Object, e As EventArgs) Handles txtSecondsDecPlaces.LostFocus
        Conversion.InputCrs.Coord.DegMinSecDecimalPlaces = txtSecondsDecPlaces.Text
        Conversion.OutputCrs.Coord.DegMinSecDecimalPlaces = txtSecondsDecPlaces.Text
    End Sub

    Private Sub chkDmsSymbols_CheckedChanged(sender As Object, e As EventArgs) Handles chkDmsSymbols.CheckedChanged
        If chkDmsSymbols.Checked Then
            Conversion.InputCrs.Coord.ShowDmsSymbols = True
            Conversion.OutputCrs.Coord.ShowDmsSymbols = True
        Else
            Conversion.InputCrs.Coord.ShowDmsSymbols = False
            Conversion.OutputCrs.Coord.ShowDmsSymbols = False
        End If
    End Sub



    'Private Sub rbEnterInputXYZ_CheckedChanged_1(sender As Object, e As EventArgs) Handles rbEnterInputXYZ.CheckedChanged

    'End Sub

    'Private Sub rbEnterInputLongLat_CheckedChanged_1(sender As Object, e As EventArgs) Handles rbEnterInputLongLat.CheckedChanged

    'End Sub

    'Private Sub rbEnterInputEastNorth_CheckedChanged_1(sender As Object, e As EventArgs) Handles rbEnterInputEastNorth.CheckedChanged

    'End Sub



    Private Sub rbDecDegrees_CheckedChanged(sender As Object, e As EventArgs) Handles rbDecDegrees.CheckedChanged
        UpdateOutputGeographicCoords()
    End Sub

    Private Sub rbDMS_CheckedChanged(sender As Object, e As EventArgs) Handles rbDMS.CheckedChanged
        UpdateOutputGeographicCoords()
    End Sub

    Private Sub rbInputEastNorth_CheckedChanged(sender As Object, e As EventArgs) Handles rbInputEastNorth.CheckedChanged
        If rbInputEastNorth.Checked Then UpdateInputLocationDataTableInput()
    End Sub

    Private Sub rbInputLongLat_CheckedChanged(sender As Object, e As EventArgs) Handles rbInputLongLat.CheckedChanged
        If rbInputLongLat.Checked Then UpdateInputLocationDataTableInput()
    End Sub

    Private Sub rbInputXYZ_CheckedChanged(sender As Object, e As EventArgs) Handles rbInputXYZ.CheckedChanged
        If rbInputXYZ.Checked Then UpdateInputLocationDataTableInput()
    End Sub

    Private Sub UpdateInputLocationDataTableInput()
        'Update the Datum Transformation Table Input columns: dgvInputLocations
        'This method sets the data entry columns to Read-Write with a white background.

        If rbInputEastNorth.Checked Then
            For Each Col As DataGridViewColumn In dgvInputLocations.Columns
                If Col.HeaderText = "Easting" Or Col.HeaderText = "Northing" Then
                    Col.ReadOnly = False
                    Col.DefaultCellStyle.BackColor = Color.White
                Else
                    Col.ReadOnly = True
                    Col.DefaultCellStyle.BackColor = Color.WhiteSmoke
                End If
            Next
        ElseIf rbInputLongLat.Checked Then
            For Each Col As DataGridViewColumn In dgvInputLocations.Columns
                If Col.HeaderText = "Longitude" Or Col.HeaderText = "Latitude" Or Col.HeaderText = "Ellipsoidal Height" Then
                    Col.ReadOnly = False
                    Col.DefaultCellStyle.BackColor = Color.White
                Else
                    Col.ReadOnly = True
                    Col.DefaultCellStyle.BackColor = Color.WhiteSmoke
                End If
            Next
        ElseIf rbInputXYZ.Checked Then
            For Each Col As DataGridViewColumn In dgvInputLocations.Columns
                If Col.HeaderText = "X" Or Col.HeaderText = "Y" Or Col.HeaderText = "Z" Then
                    Col.ReadOnly = False
                    Col.DefaultCellStyle.BackColor = Color.White
                Else
                    Col.ReadOnly = True
                    Col.DefaultCellStyle.BackColor = Color.WhiteSmoke
                End If
            Next
        End If
    End Sub

    Private Sub UpdateOutputLocationDataTableInput()
        'Update the Datum Transformation Table Output columns: dgvOutputLocations
        'This method sets the data entry columns to Read-Write with a white background.

        If rbOutputEastNorth.Checked Then
            For Each Col As DataGridViewColumn In dgvOutputLocations.Columns
                If Col.HeaderText = "Easting" Or Col.HeaderText = "Northing" Then
                    Col.ReadOnly = False
                    Col.DefaultCellStyle.BackColor = Color.White
                Else
                    Col.ReadOnly = True
                    Col.DefaultCellStyle.BackColor = Color.WhiteSmoke
                End If
            Next
        ElseIf rbOutputLongLat.Checked Then
            For Each Col As DataGridViewColumn In dgvOutputLocations.Columns
                If Col.HeaderText = "Longitude" Or Col.HeaderText = "Latitude" Or Col.HeaderText = "Ellipsoidal Height" Then
                    Col.ReadOnly = False
                    Col.DefaultCellStyle.BackColor = Color.White
                Else
                    Col.ReadOnly = True
                    Col.DefaultCellStyle.BackColor = Color.WhiteSmoke
                End If
            Next
        ElseIf rbOutputXYZ.Checked Then
            For Each Col As DataGridViewColumn In dgvOutputLocations.Columns
                If Col.HeaderText = "X" Or Col.HeaderText = "Y" Or Col.HeaderText = "Z" Then
                    Col.ReadOnly = False
                    Col.DefaultCellStyle.BackColor = Color.White
                Else
                    Col.ReadOnly = True
                    Col.DefaultCellStyle.BackColor = Color.WhiteSmoke
                End If
            Next
        End If
    End Sub

    Private Sub dgvInputLocations_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvInputLocations.CellContentClick

    End Sub

    Private Sub dgvInputLocations_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvInputLocations.CellClick

        Dim ColHeader As String = dgvInputLocations.Columns(e.ColumnIndex).HeaderText

        Select Case ColHeader
            Case "Easting"

            Case "Northing"

            Case "Longitude"
                Conversion.Angle.SetAngle(dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                Conversion.Angle.DecDegreeDecPlaces = txtInputDegreeDecPlaces.Text
                Conversion.Angle.SecondsDecPlaces = txtInputSecondsDecPlaces.Text
                txtInputDecDegree.Text = Conversion.Angle.FormattedDecimalDegrees
                If Conversion.Angle.Sign = clsAngle.AngleSign.Negative Then txtInputAngleSign.Text = "-" Else txtInputAngleSign.Text = "+"
                txtInputAngleDegrees.Text = Conversion.Angle.Degrees
                txtInputAngleMinutes.Text = Conversion.Angle.Minutes
                txtInputAngleSeconds.Text = Conversion.Angle.FormattedSeconds

            Case "Latitude"
                Conversion.Angle.SetAngle(dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                Conversion.Angle.DecDegreeDecPlaces = txtInputDegreeDecPlaces.Text
                Conversion.Angle.SecondsDecPlaces = txtInputSecondsDecPlaces.Text
                txtInputDecDegree.Text = Conversion.Angle.FormattedDecimalDegrees
                If Conversion.Angle.Sign = clsAngle.AngleSign.Negative Then txtInputAngleSign.Text = "-" Else txtInputAngleSign.Text = "+"
                txtInputAngleDegrees.Text = Conversion.Angle.Degrees
                txtInputAngleMinutes.Text = Conversion.Angle.Minutes
                txtInputAngleSeconds.Text = Conversion.Angle.FormattedSeconds

            Case "Ellipsoidal Height"

            Case "X"

            Case "Y"

            Case "Z"

        End Select
    End Sub

    Private Sub txtOutputCrsProjMethodFormula_TextChanged(sender As Object, e As EventArgs) Handles txtOutputCrsProjMethodFormula.TextChanged

    End Sub

    Private Sub txtInputProjFormat_TextChanged(sender As Object, e As EventArgs) Handles txtInputProjFormat.TextChanged

    End Sub

    Private Sub txtInputProjFormat_LostFocus(sender As Object, e As EventArgs) Handles txtInputProjFormat.LostFocus

        dgvInputLocations.Columns(0).DefaultCellStyle.Format = txtInputProjFormat.Text
        dgvInputLocations.Columns(1).DefaultCellStyle.Format = txtInputProjFormat.Text
    End Sub

    Private Sub txtInputCartFormat_TextChanged(sender As Object, e As EventArgs) Handles txtInputCartFormat.TextChanged

    End Sub

    Private Sub txtInputCartFormat_LostFocus(sender As Object, e As EventArgs) Handles txtInputCartFormat.LostFocus
        dgvInputLocations.Columns(5).DefaultCellStyle.Format = txtInputCartFormat.Text
        dgvInputLocations.Columns(6).DefaultCellStyle.Format = txtInputCartFormat.Text
        dgvInputLocations.Columns(7).DefaultCellStyle.Format = txtInputCartFormat.Text
    End Sub

    Private Sub txtInputHeightFormat_TextChanged(sender As Object, e As EventArgs) Handles txtInputHeightFormat.TextChanged

    End Sub

    Private Sub txtInputHeightFormat_LostFocus(sender As Object, e As EventArgs) Handles txtInputHeightFormat.LostFocus
        dgvInputLocations.Columns(4).DefaultCellStyle.Format = txtInputHeightFormat.Text
    End Sub

    Private Sub txtInputDegreeDecPlaces_TextChanged(sender As Object, e As EventArgs) Handles txtInputDegreeDecPlaces.TextChanged

    End Sub

    Private Sub txtInputDegreeDecPlaces_LostFocus(sender As Object, e As EventArgs) Handles txtInputDegreeDecPlaces.LostFocus

        If rbInputDecDegrees.Checked Then
            dgvInputLocations.Columns(2).DefaultCellStyle.Format = "F" & txtInputDegreeDecPlaces.Text
            dgvInputLocations.Columns(3).DefaultCellStyle.Format = "F" & txtInputDegreeDecPlaces.Text
        End If
    End Sub

    Private Sub txtInputSecondsDecPlaces_TextChanged(sender As Object, e As EventArgs) Handles txtInputSecondsDecPlaces.TextChanged

    End Sub

    Private Sub txtInputSecondsDecPlaces_LostFocus(sender As Object, e As EventArgs) Handles txtInputSecondsDecPlaces.LostFocus
        Conversion.InputCrs.Coord.DegMinSecDecimalPlaces = txtInputSecondsDecPlaces.Text
    End Sub

    Private Sub chkInputDmsSymbols_CheckedChanged(sender As Object, e As EventArgs) Handles chkInputDmsSymbols.CheckedChanged
        If chkInputDmsSymbols.Checked Then Conversion.InputCrs.Coord.ShowDmsSymbols = True Else Conversion.InputCrs.Coord.ShowDmsSymbols = False
    End Sub

    Private Sub dgvOutputLocations_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOutputLocations.CellContentClick

    End Sub

    Private Sub dgvOutputLocations_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOutputLocations.CellClick

        Dim ColHeader As String = dgvOutputLocations.Columns(e.ColumnIndex).HeaderText

        Select Case ColHeader
            Case "Easting"

            Case "Northing"

            Case "Longitude"
                Conversion.Angle.SetAngle(dgvOutputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                If txtOutputDegreeDecPlaces.Text.Trim = "" Then txtOutputDegreeDecPlaces.Text = "7"
                Conversion.Angle.DecDegreeDecPlaces = txtOutputDegreeDecPlaces.Text
                If txtOutputSecondsDecPlaces.Text.Trim = "" Then txtOutputSecondsDecPlaces.Text = "4"
                Conversion.Angle.SecondsDecPlaces = txtOutputSecondsDecPlaces.Text
                txtOutputDecDegree.Text = Conversion.Angle.FormattedDecimalDegrees
                If Conversion.Angle.Sign = clsAngle.AngleSign.Negative Then txtOutputAngleSign.Text = "-" Else txtOutputAngleSign.Text = "+"
                txtOutputAngleDegrees.Text = Conversion.Angle.Degrees
                txtOutputAngleMinutes.Text = Conversion.Angle.Minutes
                txtOutputAngleSeconds.Text = Conversion.Angle.FormattedSeconds

            Case "Latitude"
                Conversion.Angle.SetAngle(dgvOutputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                If txtOutputDegreeDecPlaces.Text.Trim = "" Then txtOutputDegreeDecPlaces.Text = "7"
                Conversion.Angle.DecDegreeDecPlaces = txtOutputDegreeDecPlaces.Text
                If txtOutputSecondsDecPlaces.Text.Trim = "" Then txtOutputSecondsDecPlaces.Text = "4"
                Conversion.Angle.SecondsDecPlaces = txtOutputSecondsDecPlaces.Text
                txtOutputDecDegree.Text = Conversion.Angle.FormattedDecimalDegrees
                If Conversion.Angle.Sign = clsAngle.AngleSign.Negative Then txtOutputAngleSign.Text = "-" Else txtOutputAngleSign.Text = "+"
                txtOutputAngleDegrees.Text = Conversion.Angle.Degrees
                txtOutputAngleMinutes.Text = Conversion.Angle.Minutes
                txtOutputAngleSeconds.Text = Conversion.Angle.FormattedSeconds

            Case "Ellipsoidal Height"

            Case "X"

            Case "Y"

            Case "Z"

        End Select

    End Sub

    Private Sub dgvOutputLocations_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOutputLocations.CellEndEdit
        'The cell has been edited - update the coordinates.

        Dim ColHeader As String = dgvOutputLocations.Columns(e.ColumnIndex).HeaderText
        Select Case ColHeader
            Case "Easting"
                If rbOutputEastNorth.Checked Then 'This is a valid entry coordinate.
                    'Check that the Northing value is valid:
                    'If dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).ToString.Trim = "" Then  'There is no Northing value
                    If dgvOutputLocations.Rows(e.RowIndex).Cells(1).ToString.Trim = "" Then  'There is no Northing value

                    Else
                        'Conversion.InputCrs.Coord.SetEasting(dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.None) 'Update the new Input Easting value
                        Conversion.OutputCrs.Coord.SetEasting(dgvOutputLocations.Rows(e.RowIndex).Cells(0).Value, Coordinate.UpdateMode.None) 'Read the Output Easting value
                        'Conversion.InputCrs.Coord.SetNorthing(dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value, Coordinate.UpdateMode.All) 'Read the Input Northing Update all the Input coordinate types
                        Conversion.OutputCrs.Coord.SetNorthing(dgvOutputLocations.Rows(e.RowIndex).Cells(1).Value, Coordinate.UpdateMode.All) 'Read the Output Northing Update all the Output coordinate types
                        If rbOutputDecDegrees.Checked Then
                            dgvOutputLocations.Rows(e.RowIndex).Cells(2).Value = Conversion.OutputCrs.Coord.Longitude
                            dgvOutputLocations.Rows(e.RowIndex).Cells(3).Value = Conversion.OutputCrs.Coord.Latitude
                        Else
                            dgvOutputLocations.Rows(e.RowIndex).Cells(2).Value = Conversion.OutputCrs.Coord.LongitudeDMS
                            dgvOutputLocations.Rows(e.RowIndex).Cells(3).Value = Conversion.OutputCrs.Coord.LatitudeDMS
                        End If
                        dgvOutputLocations.Rows(e.RowIndex).Cells(4).Value = Conversion.OutputCrs.Coord.EllipsoidalHeight
                        dgvOutputLocations.Rows(e.RowIndex).Cells(5).Value = Conversion.OutputCrs.Coord.X
                        dgvOutputLocations.Rows(e.RowIndex).Cells(6).Value = Conversion.OutputCrs.Coord.Y
                        dgvOutputLocations.Rows(e.RowIndex).Cells(7).Value = Conversion.OutputCrs.Coord.Z
                        dgvOutputLocations.AutoResizeColumns()
                    End If
                End If
            Case "Northing"
                If rbOutputEastNorth.Checked Then 'This is a valid entry coordinate.
                    'Check that the Easting value is valid:
                    'If dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).ToString.Trim = "" Then  'There is no Easting value
                    If dgvOutputLocations.Rows(e.RowIndex).Cells(0).ToString.Trim = "" Then  'There is no Easting value

                    Else
                        'Conversion.InputCrs.Coord.SetEasting(dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.None) 'Read the Input Easting value
                        Conversion.OutputCrs.Coord.SetEasting(dgvOutputLocations.Rows(e.RowIndex).Cells(0).Value, Coordinate.UpdateMode.None) 'Read the Output Easting value
                        'Conversion.InputCrs.Coord.SetNorthing(dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value, Coordinate.UpdateMode.All) 'Read the Input Northing Update all the Input coordinate types
                        Conversion.OutputCrs.Coord.SetNorthing(dgvOutputLocations.Rows(e.RowIndex).Cells(1).Value, Coordinate.UpdateMode.All) 'Read the Output Northing and Update all the Output coordinate types
                        If rbOutputDecDegrees.Checked Then
                            dgvOutputLocations.Rows(e.RowIndex).Cells(2).Value = Conversion.OutputCrs.Coord.Longitude
                            dgvOutputLocations.Rows(e.RowIndex).Cells(3).Value = Conversion.OutputCrs.Coord.Latitude
                        Else
                            dgvOutputLocations.Rows(e.RowIndex).Cells(2).Value = Conversion.OutputCrs.Coord.LongitudeDMS
                            dgvOutputLocations.Rows(e.RowIndex).Cells(3).Value = Conversion.OutputCrs.Coord.LatitudeDMS
                        End If
                        dgvOutputLocations.Rows(e.RowIndex).Cells(4).Value = Conversion.OutputCrs.Coord.EllipsoidalHeight
                        dgvOutputLocations.Rows(e.RowIndex).Cells(5).Value = Conversion.OutputCrs.Coord.X
                        dgvOutputLocations.Rows(e.RowIndex).Cells(6).Value = Conversion.OutputCrs.Coord.Y
                        dgvOutputLocations.Rows(e.RowIndex).Cells(7).Value = Conversion.OutputCrs.Coord.Z
                        dgvOutputLocations.AutoResizeColumns()
                    End If
                End If

            Case "Longitude"
                If rbOutputLongLat.Checked Then 'This is a valid entry coordinate.
                    'Check that the Ellipsoidal Height value is valid:
                    If dgvOutputLocations.Rows(e.RowIndex).Cells(4).ToString.Trim = "" Then  'There is no Ellipsoidal Height value.
                        dgvOutputLocations.Rows(e.RowIndex).Cells(4).Value = 0 'Set the Ellipsoidal Height to the default value of 0.
                    End If
                    'Check that the Latitude value is valid:
                    If dgvOutputLocations.Rows(e.RowIndex).Cells(3).ToString.Trim = "" Then  'There is no Latitude value

                    Else
                        Conversion.OutputCrs.Coord.SetLatitude(dgvOutputLocations.Rows(e.RowIndex).Cells(3).Value, Coordinate.UpdateMode.None) 'Read the Output Latitude value
                        Conversion.OutputCrs.Coord.SetEllipsoidalHeight(dgvOutputLocations.Rows(e.RowIndex).Cells(4).Value, Coordinate.UpdateMode.None) 'Read the Output Ellipsoidal Height value
                        Conversion.OutputCrs.Coord.SetLongitude(dgvOutputLocations.Rows(e.RowIndex).Cells(2).Value, Coordinate.UpdateMode.All) 'Read the Output Longitude value and Update all the Output coordinate types
                        dgvOutputLocations.Rows(e.RowIndex).Cells(0).Value = Conversion.OutputCrs.Coord.Easting
                        dgvOutputLocations.Rows(e.RowIndex).Cells(1).Value = Conversion.OutputCrs.Coord.Northing
                        dgvOutputLocations.Rows(e.RowIndex).Cells(5).Value = Conversion.OutputCrs.Coord.X
                        dgvOutputLocations.Rows(e.RowIndex).Cells(6).Value = Conversion.OutputCrs.Coord.Y
                        dgvOutputLocations.Rows(e.RowIndex).Cells(7).Value = Conversion.OutputCrs.Coord.Z
                        dgvOutputLocations.AutoResizeColumns()
                    End If
                End If

            Case "Latitude"
                If rbOutputLongLat.Checked Then 'This is a valid entry coordinate.
                    'Check that the Ellipsoidal Height value is valid:
                    If dgvOutputLocations.Rows(e.RowIndex).Cells(4).ToString.Trim = "" Then  'There is no Ellipsoidal Height value.
                        dgvOutputLocations.Rows(e.RowIndex).Cells(4).Value = 0 'Set the Ellipsoidal Height to the default value of 0.
                    End If
                    'Check that the Longitude value is valid:
                    If dgvOutputLocations.Rows(e.RowIndex).Cells(2).ToString.Trim = "" Then  'There is no Longitude value

                    Else
                        Conversion.OutputCrs.Coord.SetLatitude(dgvOutputLocations.Rows(e.RowIndex).Cells(3).Value, Coordinate.UpdateMode.None) 'Read the Output Latitude value
                        Conversion.OutputCrs.Coord.SetEllipsoidalHeight(dgvOutputLocations.Rows(e.RowIndex).Cells(4).Value, Coordinate.UpdateMode.None) 'Read the Output Ellipsoidal Height value
                        Conversion.OutputCrs.Coord.SetLongitude(dgvOutputLocations.Rows(e.RowIndex).Cells(2).Value, Coordinate.UpdateMode.All) 'Read the Output Longitude value and Update all the Output coordinate types
                        dgvOutputLocations.Rows(e.RowIndex).Cells(0).Value = Conversion.OutputCrs.Coord.Easting
                        dgvOutputLocations.Rows(e.RowIndex).Cells(1).Value = Conversion.OutputCrs.Coord.Northing
                        dgvOutputLocations.Rows(e.RowIndex).Cells(5).Value = Conversion.OutputCrs.Coord.X
                        dgvOutputLocations.Rows(e.RowIndex).Cells(6).Value = Conversion.OutputCrs.Coord.Y
                        dgvOutputLocations.Rows(e.RowIndex).Cells(7).Value = Conversion.OutputCrs.Coord.Z
                        dgvOutputLocations.AutoResizeColumns()
                    End If
                End If

            Case "Ellipsoidal Height"
                If rbOutputLongLat.Checked Then 'This is a valid entry coordinate.
                    'Check that the Longitude and Latitude values are valid:
                    If dgvOutputLocations.Rows(e.RowIndex).Cells(2).ToString.Trim = "" And dgvOutputLocations.Rows(e.RowIndex).Cells(3).ToString.Trim = "" Then  'There is no Longitude or Latitude value

                    Else
                        Conversion.OutputCrs.Coord.SetLatitude(dgvOutputLocations.Rows(e.RowIndex).Cells(3).Value, Coordinate.UpdateMode.None) 'Read the Output Latitude value
                        Conversion.OutputCrs.Coord.SetEllipsoidalHeight(dgvOutputLocations.Rows(e.RowIndex).Cells(4).Value, Coordinate.UpdateMode.None) 'Read the Output Ellipsoidal Height value
                        Conversion.OutputCrs.Coord.SetLongitude(dgvOutputLocations.Rows(e.RowIndex).Cells(2).Value, Coordinate.UpdateMode.All) 'Read the Output Longitude value and Update all the Output coordinate types
                        dgvOutputLocations.Rows(e.RowIndex).Cells(0).Value = Conversion.OutputCrs.Coord.Easting
                        dgvOutputLocations.Rows(e.RowIndex).Cells(1).Value = Conversion.OutputCrs.Coord.Northing
                        dgvOutputLocations.Rows(e.RowIndex).Cells(5).Value = Conversion.OutputCrs.Coord.X
                        dgvOutputLocations.Rows(e.RowIndex).Cells(6).Value = Conversion.OutputCrs.Coord.Y
                        dgvOutputLocations.Rows(e.RowIndex).Cells(7).Value = Conversion.OutputCrs.Coord.Z
                        dgvOutputLocations.AutoResizeColumns()
                    End If
                End If

            Case "X"

            Case "Y"

            Case "Z"

            Case Else

        End Select

    End Sub

    Private Sub rbOutputEastNorth_CheckedChanged(sender As Object, e As EventArgs) Handles rbOutputEastNorth.CheckedChanged
        If rbOutputEastNorth.Checked Then UpdateOutputLocationDataTableInput()
    End Sub

    Private Sub rbOutputLongLat_CheckedChanged(sender As Object, e As EventArgs) Handles rbOutputLongLat.CheckedChanged
        If rbOutputLongLat.Checked Then UpdateOutputLocationDataTableInput()
    End Sub

    Private Sub rbOutputXYZ_CheckedChanged(sender As Object, e As EventArgs) Handles rbOutputXYZ.CheckedChanged
        If rbOutputXYZ.Checked Then UpdateOutputLocationDataTableInput()
    End Sub

    Private Sub txtOutputProjFormat_TextChanged(sender As Object, e As EventArgs) Handles txtOutputProjFormat.TextChanged

    End Sub

    Private Sub txtOutputProjFormat_LostFocus(sender As Object, e As EventArgs) Handles txtOutputProjFormat.LostFocus
        dgvOutputLocations.Columns(0).DefaultCellStyle.Format = txtOutputProjFormat.Text
        dgvOutputLocations.Columns(1).DefaultCellStyle.Format = txtOutputProjFormat.Text

    End Sub

    Private Sub txtOutputCartFormat_TextChanged(sender As Object, e As EventArgs) Handles txtOutputCartFormat.TextChanged

    End Sub

    Private Sub txtOutputCartFormat_LostFocus(sender As Object, e As EventArgs) Handles txtOutputCartFormat.LostFocus
        dgvOutputLocations.Columns(5).DefaultCellStyle.Format = txtOutputCartFormat.Text
        dgvOutputLocations.Columns(6).DefaultCellStyle.Format = txtOutputCartFormat.Text
        dgvOutputLocations.Columns(7).DefaultCellStyle.Format = txtOutputCartFormat.Text

    End Sub

    Private Sub rbInputDecDegrees_CheckedChanged(sender As Object, e As EventArgs) Handles rbInputDecDegrees.CheckedChanged

    End Sub

    Private Sub rbInputDMS_CheckedChanged(sender As Object, e As EventArgs) Handles rbInputDMS.CheckedChanged

    End Sub

    Private Sub txtOutputDegreeDecPlaces_TextChanged(sender As Object, e As EventArgs) Handles txtOutputDegreeDecPlaces.TextChanged

    End Sub

    Private Sub txtOutputDegreeDecPlaces_LostFocus(sender As Object, e As EventArgs) Handles txtOutputDegreeDecPlaces.LostFocus
        If rbOutputDecDegrees.Checked Then
            dgvOutputLocations.Columns(2).DefaultCellStyle.Format = "F" & txtOutputDegreeDecPlaces.Text
            dgvOutputLocations.Columns(3).DefaultCellStyle.Format = "F" & txtOutputDegreeDecPlaces.Text
        End If

    End Sub

    Private Sub txtOutputHeightFormat_TextChanged(sender As Object, e As EventArgs) Handles txtOutputHeightFormat.TextChanged

    End Sub

    Private Sub txtOutputHeightFormat_LostFocus(sender As Object, e As EventArgs) Handles txtOutputHeightFormat.LostFocus
        dgvOutputLocations.Columns(4).DefaultCellStyle.Format = txtOutputHeightFormat.Text
    End Sub

    Private Sub txtOutputSecondsDecPlaces_TextChanged(sender As Object, e As EventArgs) Handles txtOutputSecondsDecPlaces.TextChanged

    End Sub

    Private Sub txtOutputSecondsDecPlaces_LostFocus(sender As Object, e As EventArgs) Handles txtOutputSecondsDecPlaces.LostFocus
        Conversion.OutputCrs.Coord.DegMinSecDecimalPlaces = txtOutputSecondsDecPlaces.Text
    End Sub

    Private Sub chkOutputDmsSymbols_CheckedChanged(sender As Object, e As EventArgs) Handles chkOutputDmsSymbols.CheckedChanged
        If chkOutputDmsSymbols.Checked Then Conversion.OutputCrs.Coord.ShowDmsSymbols = True Else Conversion.OutputCrs.Coord.ShowDmsSymbols = False

    End Sub

    Private Sub udPointFontSize_ValueChanged(sender As Object, e As EventArgs) Handles udPointFontSize.ValueChanged

        Dim NewFont As New Font(dgvConversion.Columns(0).DefaultCellStyle.Font.Name, udPointFontSize.Value, dgvConversion.Columns(0).DefaultCellStyle.Font.Style)
        dgvConversion.Columns(0).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(0).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(1).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(1).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(2).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(2).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(3).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(3).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(4).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(4).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(5).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(5).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(6).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(6).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(7).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(7).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(8).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(8).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(9).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(9).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(10).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(10).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(11).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(11).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(12).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(12).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(13).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(13).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(14).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(14).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(15).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(15).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(16).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(16).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(17).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(17).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(18).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(18).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(19).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(19).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(20).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(20).DefaultCellStyle.Font = NewFont
        dgvConversion.Columns(21).HeaderCell.Style.Font = NewFont
        dgvConversion.Columns(21).DefaultCellStyle.Font = NewFont
        dgvConversion.AutoResizeColumns()

    End Sub

    Private Sub btnNewConvFile_Click(sender As Object, e As EventArgs) Handles btnNewConvFile.Click
        'Create a new Coordinate Conversion file.
        Dim EntryForm As New ADVL_Utilities_Library_1.frmNewDataNameModal
        EntryForm.EntryName = "NewConvFile"
        EntryForm.Title = "New Coordinate Conversion File"
        EntryForm.FileExtension = "CoordConv"
        EntryForm.GetFileName = True
        EntryForm.GetDataName = True
        EntryForm.GetDataLabel = False
        EntryForm.GetDataDescription = True
        EntryForm.SettingsLocn = Main.Project.SettingsLocn
        EntryForm.DataLocn = Main.Project.DataLocn
        EntryForm.ApplicationName = Main.ApplicationInfo.Name
        EntryForm.RestoreFormSettings()

        If EntryForm.ShowDialog() = DialogResult.OK Then
            If txtFileName.Text.Trim = "" Then
                'There is no file to save
            Else
                If Modified Then
                    Dim Result As DialogResult = MessageBox.Show("Do you want to save the changes in the current conversion settings?", "Warning", MessageBoxButtons.YesNoCancel)
                    If Result = DialogResult.Yes Then
                        'SaveDistribModel()
                        SaveCoordConvSettings()
                    ElseIf Result = DialogResult.Cancel Then
                        Exit Sub
                    Else
                        'Contunue without saving the conversion settings.
                        Modified = False
                    End If
                End If
            End If
        End If

        ClearSettings()
        txtFileName.Text = EntryForm.FileName
        txtDataName.Text = EntryForm.DataName
        txtDescription.Text = EntryForm.DataDescription

    End Sub

    Private Sub ClearSettings()
        'Clear the Coordinate Conversion settings.
        dgvInputLocations.Rows.Clear()
        dgvOutputLocations.Rows.Clear()
        dgvConversion.Rows.Clear()
    End Sub

    Private Sub btnSaveConvFile_Click(sender As Object, e As EventArgs) Handles btnSaveConvFile.Click
        SaveCoordConvSettings()
    End Sub

    Private Sub btnSetDefault_Click(sender As Object, e As EventArgs) Handles btnSetDefault.Click
        'Set the default Datum Transformation method.

        Dim InputCrsName As String = Conversion.InputCrs.Name
        Dim InputCrsCode As Integer = Conversion.InputCrs.Code
        Dim OutputCrsName As String = Conversion.OutputCrs.Name
        Dim OutputCrsCode As Integer = Conversion.OutputCrs.Code

        Dim DatumTransType As String = Conversion.DatumTrans.Type.ToString

        Dim DatumTrans As System.Xml.Linq.XDocument

        If Conversion.DatumTrans.Type = clsDatumTrans.enumType.None Then
            'No datum transformation required.
        Else
            If Conversion.DatumTrans.Type = clsDatumTrans.enumType.Direct Then
                Dim DatumTransName As String = Conversion.DatumTrans.DirectCoordOp.Name
                Dim DatumTransCode As Integer = Conversion.DatumTrans.DirectCoordOp.Code
                Dim DatumCodeApplyReverse As Boolean = Conversion.DatumTrans.DirectMethodApplyReverse

                DatumTrans = <?xml version="1.0" encoding="utf-8"?>
                             <DatumTransformationDefaultList>
                                 <DatumTrans>
                                     <InputCrsName><%= InputCrsName %></InputCrsName>
                                     <InputCrsCode><%= InputCrsCode %></InputCrsCode>
                                     <OutputCrsName><%= OutputCrsName %></OutputCrsName>
                                     <OutputCrsCode><%= OutputCrsCode %></OutputCrsCode>
                                     <Type><%= DatumTransType %></Type>
                                     <Name><%= DatumTransName %></Name>
                                     <Code><%= DatumTransCode %></Code>
                                     <ApplyReverse><%= DatumCodeApplyReverse %></ApplyReverse>
                                 </DatumTrans>
                             </DatumTransformationDefaultList>
            ElseIf Conversion.DatumTrans.Type = clsDatumTrans.enumType.ViaWgs84 Then
                Dim InputToWgs84Name As String = Conversion.DatumTrans.InputToWgs84CoordOp.Name
                Dim InputToWgs84Code As Integer = Conversion.DatumTrans.InputToWgs84CoordOp.Code
                Dim InputToWgs84ApplyReverse As Boolean = Conversion.DatumTrans.InputToWgs84MethodApplyReverse
                Dim Wgs84ToOutputName As String = Conversion.DatumTrans.Wgs84ToOutputCoordOp.Name
                Dim Wgs84ToOutputCode As Integer = Conversion.DatumTrans.Wgs84ToOutputCoordOp.Code
                Dim Wgs84ToOutputApplyReverse As Boolean = Conversion.DatumTrans.Wgs84ToOutputMethodApplyReverse
                DatumTrans = <?xml version="1.0" encoding="utf-8"?>
                             <DatumTransformationDefaultList>
                                 <DatumTrans>
                                     <InputCrsName><%= InputCrsName %></InputCrsName>
                                     <InputCrsCode><%= InputCrsCode %></InputCrsCode>
                                     <OutputCrsName><%= OutputCrsName %></OutputCrsName>
                                     <OutputCrsCode><%= OutputCrsCode %></OutputCrsCode>
                                     <Type><%= DatumTransType %></Type>
                                     <InputToWgs84Name><%= InputToWgs84Name %></InputToWgs84Name>
                                     <InputToWgs84Code><%= InputToWgs84Code %></InputToWgs84Code>
                                     <InputToWgs84ApplyReverse><%= InputToWgs84ApplyReverse %></InputToWgs84ApplyReverse>
                                     <Wgs84ToOutputName><%= Wgs84ToOutputName %></Wgs84ToOutputName>
                                     <Wgs84ToOutputCode><%= Wgs84ToOutputCode %></Wgs84ToOutputCode>
                                     <Wgs84ToOutputApplyReverse><%= Wgs84ToOutputApplyReverse %></Wgs84ToOutputApplyReverse>
                                 </DatumTrans>
                             </DatumTransformationDefaultList>
            End If

            If Main.Project.DataFileExists("DatumTransformationDefaultList.xml") Then
                Dim DatumTransDefaults As System.Xml.Linq.XDocument
                Main.Project.ReadXmlData("DatumTransformationDefaultList.xml", DatumTransDefaults)
                Dim DefaultDatumTrans = From item In DatumTransDefaults.<DatumTransformationDefaultList>.<DatumTrans> Where item.<InputCrsCode>.Value = InputCrsCode And item.<OutputCrsCode>.Value = OutputCrsCode
                If DefaultDatumTrans.Count > 0 Then
                    DefaultDatumTrans(0).ReplaceWith(DatumTrans.<DatumTransformationDefaultList>.<DatumTrans>)
                Else
                    DatumTransDefaults.<DatumTransformationDefaultList>.Nodes.First.AddAfterSelf(DatumTrans.<DatumTransformationDefaultList>.<DatumTrans>)
                End If
                Main.Project.SaveXmlData("DatumTransformationDefaultList.xml", DatumTransDefaults)
            Else
                Main.Project.SaveXmlData("DatumTransformationDefaultList.xml", DatumTrans)
            End If
        End If
    End Sub

    'Private Sub FindDefaultDatumTrans()
    Private Function FindDefaultDatumTrans() As Boolean
        'Search the Datum Transformation Default list for the Default transformation operation(s).
        'Return True if the Default was found.

        If Main.Project.DataFileExists("DatumTransformationDefaultList.xml") Then
            Dim DatumTransDefaults As System.Xml.Linq.XDocument
            Main.Project.ReadXmlData("DatumTransformationDefaultList.xml", DatumTransDefaults)
            Dim InputCrsName As String = Conversion.InputCrs.Name
            Dim InputCrsCode As Integer = Conversion.InputCrs.Code
            Dim OutputCrsName As String = Conversion.OutputCrs.Name
            Dim OutputCrsCode As Integer = Conversion.OutputCrs.Code
            Dim DefaultDatumTrans = From item In DatumTransDefaults.<DatumTransformationDefaultList>.<DatumTrans> Where item.<InputCrsCode>.Value = InputCrsCode And item.<OutputCrsCode>.Value = OutputCrsCode
            If DefaultDatumTrans.Count > 0 Then
                'If DefaultDatumTrans(0).<DatumTransformationDefaultList>.<DatumTrans>.<Type>.Value = "Direct" Then
                If DefaultDatumTrans(0).<Type>.Value = "Direct" Then
                    txtDefDatumTransType.Text = "Direct"
                    'txtDefDatumTransOpName1.Text = DefaultDatumTrans(0).<DatumTransformationDefaultList>.<DatumTrans>.<Name>.Value
                    txtDefDatumTransOpName1.Text = DefaultDatumTrans(0).<Name>.Value
                    txtDefDatumTransOpCode1.Text = DefaultDatumTrans(0).<Code>.Value
                    txtDefDatumTransOpName2.Text = ""
                    txtDefDatumTransOpCode2.Text = ""
                    Return True
                ElseIf DefaultDatumTrans(0).<Type>.Value = "ViaWgs84" Then
                    txtDefDatumTransType.Text = "Via WGS 84"
                    txtDefDatumTransOpName1.Text = DefaultDatumTrans(0).<InputToWgs84Name>.Value
                    txtDefDatumTransOpCode1.Text = DefaultDatumTrans(0).<InputToWgs84Code>.Value
                    txtDefDatumTransOpName2.Text = DefaultDatumTrans(0).<Wgs84ToOutputName>.Value
                    txtDefDatumTransOpCode2.Text = DefaultDatumTrans(0).<Wgs84ToOutputCode>.Value
                    Return True
                ElseIf DefaultDatumTrans(0).<Type>.Value = "None" Then
                    txtDefDatumTransType.Text = ""
                    txtDefDatumTransOpName1.Text = ""
                    txtDefDatumTransOpCode1.Text = ""
                    txtDefDatumTransOpName2.Text = ""
                    txtDefDatumTransOpCode2.Text = ""
                    Return True 'No Default required!
                Else
                    Main.Message.AddWarning("Unknown Datum Transformation type: " & DefaultDatumTrans(0).<Type>.Value & vbCrLf)
                    txtDefDatumTransType.Text = ""
                    txtDefDatumTransOpName1.Text = ""
                    txtDefDatumTransOpCode1.Text = ""
                    txtDefDatumTransOpName2.Text = ""
                    txtDefDatumTransOpCode2.Text = ""
                    Return False
                End If
            Else 'A Default datum transformation was not found.
                txtDefDatumTransType.Text = ""
                txtDefDatumTransOpName1.Text = ""
                txtDefDatumTransOpCode1.Text = ""
                txtDefDatumTransOpName2.Text = ""
                txtDefDatumTransOpCode2.Text = ""
                Return False
            End If
        Else 'The Default datum transformation list was not found.
            txtDefDatumTransType.Text = ""
            txtDefDatumTransOpName1.Text = ""
            txtDefDatumTransOpCode1.Text = ""
            txtDefDatumTransOpName2.Text = ""
            txtDefDatumTransOpCode2.Text = ""
            Return False
        End If

    End Function

    Private Sub FindBestDatumTransformation()
        'Find the best datum tranformation.
        'The Direct datum transformation operations will be searched first.
        'Deprecated operations will be ignored.
        'Only operations using method code 9607 (Coordinate Frame Rotation (geog 2D domain) will be searched.
        'The operation with the best accuracy will be selected.
        'If a direct operation is not found then transformations via WGS 84 will be searched.


        'Search the direct operations for the most accurate datum transformation:
        Dim DirectOpFound As Boolean = False
        Dim DirectOpCode As Integer
        Dim ApplyReverse As Boolean = False
        Dim Accuracy As Single = Single.MaxValue
        For Each Row As DataGridViewRow In dgvDirectDTOps.Rows
            If Row.Cells(13).Value = 9607 Then 'This is a Coordinate Frame Rotation (geog 2D domain)
                DirectOpFound = True
                If Row.Cells(3).Value < Accuracy Then
                    DirectOpCode = Row.Cells(2).Value
                    ApplyReverse = Row.Cells(12).Value
                    Accuracy = Row.Cells(3).Value
                End If
            End If
        Next

        If DirectOpFound Then 'Us the Direct Datum Transformation operation.
            SelectDirectDatumTransOp(DirectOpCode, ApplyReverse)
        Else 'Search the Via WGS 84 operations for the most accurate datum transformations
            Dim InputToWgs84OpFound As Boolean = False
            Dim InputToWgs84OpCode As Integer
            Dim InputToWgs84ApplyReverse As Boolean = False
            Dim InputToWgs84OpAccuracy As Single = Single.MaxValue
            For Each Row As DataGridViewRow In dgvInputToWgs84DTOps.Rows
                If Row.Cells(13).Value = 9607 Then 'This is a Coordinate Frame Rotation (geog 2D domain)
                    InputToWgs84OpFound = True
                    If Row.Cells(3).Value < InputToWgs84OpAccuracy Then
                        InputToWgs84OpCode = Row.Cells(2).Value
                        InputToWgs84ApplyReverse = Row.Cells(12).Value
                        InputToWgs84OpAccuracy = Row.Cells(3).Value
                    End If
                End If
            Next

            Dim Wgs84ToOutputOpFound As Boolean = False
            Dim Wgs84ToOutputOpCode As Integer
            Dim Wgs84ToOutputApplyReverse As Boolean = False
            Dim Wgs84ToOutputOpAccuracy As Single = Single.MaxValue
            For Each Row As DataGridViewRow In dgvWgs84ToOutputDTOps.Rows
                If Row.Cells(13).Value = 9607 Then 'This is a Coordinate Frame Rotation (geog 2D domain)
                    Wgs84ToOutputOpFound = True
                    If Row.Cells(3).Value < Wgs84ToOutputOpAccuracy Then
                        Wgs84ToOutputOpCode = Row.Cells(2).Value
                        Wgs84ToOutputApplyReverse = Row.Cells(12).Value
                        Wgs84ToOutputOpAccuracy = Row.Cells(3).Value
                    End If
                End If
            Next

            If InputToWgs84OpFound And Wgs84ToOutputOpFound Then
                SelectViaWgs84DatumTransOp(InputToWgs84OpCode, InputToWgs84ApplyReverse, Wgs84ToOutputOpCode, Wgs84ToOutputApplyReverse)
            Else
                Main.Message.AddWarning("A suitable datum transformation operation could not be found." & vbCrLf)
            End If
        End If
    End Sub

    Private Sub SelectDirectDatumTransOp(OpCode As Integer, ApplyReverse As Boolean)
        'Select the Datum Transformation Operation with code OpCode.
        SelectDirectTransOpCode(OpCode)
        Conversion.DatumTrans.GetDirectDatumTransCoordOp(OpCode)
        Conversion.DatumTrans.DirectMethodApplyReverse = ApplyReverse
    End Sub

    Private Sub SelectViaWgs84DatumTransOp(InputToWgs84OpCode As Integer, InputToWgs84ApplyReverse As Boolean, Wgs84ToOutputOpCode As Integer, Wgs84ToOutputApplyReverse As Boolean)
        'Select the Via WGS 84 Datum Transformations with codes InputToWgs84OpCode and Wgs84ToOutputOpCode.
        SelectInputToWgs84TransOpCode(InputToWgs84OpCode)
        SelectWgs84ToOutputTransOpCode(Wgs84ToOutputOpCode)
        Conversion.DatumTrans.GetInputToWgs84DatumTransCoordOp(InputToWgs84OpCode)
        Conversion.DatumTrans.InputToWgs84MethodApplyReverse = InputToWgs84ApplyReverse
        Conversion.DatumTrans.GetWgs84ToOutputDatumTransCoordOp(Wgs84ToOutputOpCode)
        Conversion.DatumTrans.Wgs84ToOutputMethodApplyReverse = Wgs84ToOutputApplyReverse
    End Sub

    Private Sub btnFormatHelp3_Click(sender As Object, e As EventArgs) Handles btnFormatHelp3.Click
        'Show Format information.
        MessageBox.Show("Format string examples:" & vbCrLf & "N4 - Number displayed with thousands separator and 4 decimal places" & vbCrLf & "F4 - Number displayed with 4 decimal places.", "Number Formatting")
    End Sub

    Private Sub SaveCoordConvSettings()
        'Save the Coordinate Conversion settings.

        Dim FileName As String = txtFileName.Text.Trim

        'Check if a file name has been specified:
        If FileName = "" Then
            Main.Message.AddWarning("Please enter a file name." & vbCrLf)
            Exit Sub
        End If

        'Check the file name extension:
        If LCase(FileName).EndsWith(".coordconv") Then
            FileName = IO.Path.GetFileNameWithoutExtension(FileName) & ".CoordConv"
        ElseIf FileName.Contains(".") Then
            Main.Message.AddWarning("Unknown file extension: " & IO.Path.GetExtension(FileName) & vbCrLf)
            Exit Sub
        Else
            FileName = FileName & ".CoordConv"
        End If

        txtFileName.Text = FileName

        dgvInputLocations.AllowUserToAddRows = False 'Otherwise a blank row is added to the saved InputLocationValues
        dgvOutputLocations.AllowUserToAddRows = False
        dgvConversion.AllowUserToAddRows = False

        'Datum Transformation:
        Dim EntryCoordType As String
        If rbEnterInputEastNorth.Checked Then
            EntryCoordType = "Projected"
        ElseIf rbEnterInputLongLat.Checked Then
            EntryCoordType = "Geographic"
        ElseIf rbEnterInputXYZ.Checked Then
            EntryCoordType = "Cartesian"
        Else
            EntryCoordType = "Geographic"
        End If

        Dim InputEntryCoordType As String
        If rbInputEastNorth.Checked Then
            InputEntryCoordType = "Projected"
        ElseIf rbInputLongLat.Checked Then
            InputEntryCoordType = "Geographic"
        ElseIf rbInputXYZ.Checked Then
            InputEntryCoordType = "Cartesian"
        Else
            InputEntryCoordType = "Geographic"
        End If

        Dim InputDegreeDisplay As String
        If rbInputDecDegrees.Checked Then
            InputDegreeDisplay = "DecimalDegrees"
        ElseIf rbInputDMS.Checked Then
            InputDegreeDisplay = "Deg-Min-Sec"
        Else
            InputDegreeDisplay = "DecimalDegrees"
        End If

        Dim OutputDegreeDisplay As String
        If rbOutputDecDegrees.Checked Then
            OutputDegreeDisplay = "DecimalDegrees"
        ElseIf rbOutputDMS.Checked Then
            OutputDegreeDisplay = "Deg-Min-Sec"
        Else
            OutputDegreeDisplay = "DecimalDegrees"
        End If

        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <CoordinateConversionSettings>
                               <SettingsName><%= txtDataName.Text %></SettingsName>
                               <SettingsDescription><%= txtDescription.Text %></SettingsDescription>
                               <!--Other Settings-->
                               <SelectedTabIndex><%= TabControl1.SelectedIndex %></SelectedTabIndex>
                               <!--Input CRS Settings-->
                               <InputEntryCoordType><%= InputEntryCoordType %></InputEntryCoordType>
                               <SelectedInputTabIndex><%= TabControl2.SelectedIndex %></SelectedInputTabIndex>
                               <InputCrsQuery><%= txtInputCrsQuery.Text %></InputCrsQuery>
                               <InputQueryNameContains><%= txtFindInput.Text %></InputQueryNameContains>
                               <InputCrsCode><%= txtInputCrsCode.Text %></InputCrsCode>
                               <SplitDist1><%= SplitContainer1.SplitterDistance %></SplitDist1>
                               <SplitDist2><%= SplitContainer2.SplitterDistance %></SplitDist2>
                               <SplitDist5><%= SplitContainer5.SplitterDistance %></SplitDist5>
                               <SelInputRowNo><%= SelInputRowNo %></SelInputRowNo>
                               <!--Save Input Location values-->
                               <InputDegreeDisplay><%= InputDegreeDisplay %></InputDegreeDisplay>
                               <InputDegreesDecPlaces><%= txtInputDegreeDecPlaces.Text %></InputDegreesDecPlaces>
                               <InputSecondsDecPlaces><%= txtInputSecondsDecPlaces.Text %></InputSecondsDecPlaces>
                               <InputShowDmsSymbols><%= chkInputDmsSymbols.Checked %></InputShowDmsSymbols>
                               <InputHeightFormat><%= txtInputHeightFormat.Text %></InputHeightFormat>
                               <InputProjectedFormat><%= txtInputProjFormat.Text %></InputProjectedFormat>
                               <InputCartesianFormat><%= txtInputCartFormat.Text %></InputCartesianFormat>
                               <InputLocationValues>
                                   <%= From Row As DataGridViewRow In dgvInputLocations.Rows
                                       Select
                                         <Row>
                                             <Easting><%= Row.Cells(0).Value %></Easting>
                                             <Northing><%= Row.Cells(1).Value %></Northing>
                                             <Longitude><%= Row.Cells(2).Value %></Longitude>
                                             <Latitude><%= Row.Cells(3).Value %></Latitude>
                                             <EllipsoidalHeight><%= Row.Cells(4).Value %></EllipsoidalHeight>
                                             <X><%= Row.Cells(5).Value %></X>
                                             <Y><%= Row.Cells(6).Value %></Y>
                                             <Z><%= Row.Cells(7).Value %></Z>
                                         </Row> %>
                               </InputLocationValues>
                               <!--Output CRS Settings-->
                               <SelectedOutputTabIndex><%= TabControl3.SelectedIndex %></SelectedOutputTabIndex>
                               <OutputCrsQuery><%= txtOutputCrsQuery.Text %></OutputCrsQuery>
                               <OutputQueryNameContains><%= txtFindOutput.Text %></OutputQueryNameContains>
                               <OutputCrsCode><%= txtOutputCrsCode.Text %></OutputCrsCode>
                               <SplitDist3><%= SplitContainer3.SplitterDistance %></SplitDist3>
                               <SplitDist4><%= SplitContainer4.SplitterDistance %></SplitDist4>
                               <SplitDist6><%= SplitContainer6.SplitterDistance %></SplitDist6>
                               <SelOutputRowNo><%= SelOutputRowNo %></SelOutputRowNo>
                               <!--Save Output Location values-->
                               <OutputDegreeDisplay><%= OutputDegreeDisplay %></OutputDegreeDisplay>
                               <OutputDegreesDecPlaces><%= txtOutputDegreeDecPlaces.Text %></OutputDegreesDecPlaces>
                               <OutputSecondsDecPlaces><%= txtOutputSecondsDecPlaces.Text %></OutputSecondsDecPlaces>
                               <OutputShowDmsSymbols><%= chkOutputDmsSymbols.Checked %></OutputShowDmsSymbols>
                               <OutputHeightFormat><%= txtOutputHeightFormat.Text %></OutputHeightFormat>
                               <OutputProjectedFormat><%= txtOutputProjFormat.Text %></OutputProjectedFormat>
                               <OutputCartesianFormat><%= txtOutputCartFormat.Text %></OutputCartesianFormat>
                               <OutputLocationValues>
                                   <%= From Row As DataGridViewRow In dgvOutputLocations.Rows
                                       Select
                                         <Row>
                                             <Easting><%= Row.Cells(0).Value %></Easting>
                                             <Northing><%= Row.Cells(1).Value %></Northing>
                                             <Longitude><%= Row.Cells(2).Value %></Longitude>
                                             <Latitude><%= Row.Cells(3).Value %></Latitude>
                                             <EllipsoidalHeight><%= Row.Cells(4).Value %></EllipsoidalHeight>
                                             <X><%= Row.Cells(5).Value %></X>
                                             <Y><%= Row.Cells(6).Value %></Y>
                                             <Z><%= Row.Cells(7).Value %></Z>
                                         </Row> %>
                               </OutputLocationValues>
                               <!--Datum Transformation Settings-->
                               <SelectedDatumTransTabIndex><%= TabControl4.SelectedIndex %></SelectedDatumTransTabIndex>
                               <SelectedViaWgs84TabIndex><%= TabControl5.SelectedIndex %></SelectedViaWgs84TabIndex>
                               <SplitDist7><%= SplitContainer7.SplitterDistance %></SplitDist7>
                               <SplitDist8><%= SplitContainer8.SplitterDistance %></SplitDist8>
                               <SplitDist9><%= SplitContainer9.SplitterDistance %></SplitDist9>
                               <SplitDist10><%= SplitContainer10.SplitterDistance %></SplitDist10>
                               <SplitDist11><%= SplitContainer11.SplitterDistance %></SplitDist11>
                               <SplitDist12><%= SplitContainer12.SplitterDistance %></SplitDist12>
                               <SplitDist13><%= SplitContainer13.SplitterDistance %></SplitDist13>
                               <SplitDist14><%= SplitContainer14.SplitterDistance %></SplitDist14>
                               <SplitDist15><%= SplitContainer15.SplitterDistance %></SplitDist15>
                               <DirectCoordOpCode><%= Conversion.DatumTrans.DirectCoordOp.Code %></DirectCoordOpCode>
                               <InputToWgs84CoordOpCode><%= Conversion.DatumTrans.InputToWgs84CoordOp.Code %></InputToWgs84CoordOpCode>
                               <Wgs84ToOutputCoordOpCode><%= Conversion.DatumTrans.Wgs84ToOutputCoordOp.Code %></Wgs84ToOutputCoordOpCode>
                               <!--  Coordinate Type Conversion Settings-->
                               <DatumTransType><%= Conversion.DatumTrans.Type %></DatumTransType>
                               <!--  Datum Transformation Settings-->
                               <EntryCoordType><%= EntryCoordType %></EntryCoordType>
                               <ShowInputEastingNorthing><%= chkShowInputEastNorth.Checked %></ShowInputEastingNorthing>
                               <ShowInputLongitudeLatitude><%= chkShowInputLongLat.Checked %></ShowInputLongitudeLatitude>
                               <ShowInputXYS><%= chkShowInputXYZ.Checked %></ShowInputXYS>
                               <ShowWgs84XYZ><%= chkShowWgs84XYZ.Checked %></ShowWgs84XYZ>
                               <ShowOutputEastingNorthing><%= chkShowOutputEastNorth.Checked %></ShowOutputEastingNorthing>
                               <ShowOutputLongitudeLatitude><%= chkShowOutputLongLat.Checked %></ShowOutputLongitudeLatitude>
                               <ShowOutputXYZ><%= chkShowOutputXYZ.Checked %></ShowOutputXYZ>
                               <ShowPointNumber><%= chkShowPointNumber.Checked %></ShowPointNumber>
                               <ShowPointName><%= chkShowPointName.Checked %></ShowPointName>
                               <ShowPointDescription><%= chkShowPointDescription.Checked %></ShowPointDescription>
                               <ProjectedFormat><%= txtProjFormat.Text %></ProjectedFormat>
                               <CartesianFormat><%= txtCartFormat.Text %></CartesianFormat>
                               <ShowDegMinSec><%= rbDMS.Checked %></ShowDegMinSec>
                               <ShowDegMinSecSymbols><%= chkDmsSymbols.Checked %></ShowDegMinSecSymbols>
                               <DecDegreesDecPlaces><%= txtDegreeDecPlaces.Text %></DecDegreesDecPlaces>
                               <DmsSecondsDecPlaces><%= txtSecondsDecPlaces.Text %></DmsSecondsDecPlaces>
                               <HeightFormat><%= txtHeightFormat.Text %></HeightFormat>
                               <!--Save the Datum Transformation Input Points-->
                               <%= DatumTransInputData(EntryCoordType).<DatumTransInputData> %>
                           </CoordinateConversionSettings>

        Main.Project.SaveXmlSettings(FileName, settingsData)

        dgvInputLocations.AllowUserToAddRows = True
        dgvOutputLocations.AllowUserToAddRows = True
        dgvConversion.AllowUserToAddRows = True

    End Sub

    Private Sub btnOpenConvFile_Click(sender As Object, e As EventArgs) Handles btnOpenConvFile.Click
        'Open a Coordinate Conversion settings file.

        Dim FileName As String = Main.Project.SelectDataFile("Coordinate Conversion files", "CoordConv")
        If FileName = "" Then
            'No file has been selected.
        Else
            'OpenDistModel(FileName)
            'GenerateData()
            OpenCoordConvSettings(FileName)
        End If
    End Sub

    Private Sub OpenCoordConvSettings(FileName As String)
        'Open the Coordinate Conversion settings file named FileName.

        ClearSettings()
        Dim Settings As System.Xml.Linq.XDocument
        Main.Project.ReadXmlData(FileName, Settings)

        If IsNothing(Settings) Then 'There is no Settings XML data.
            Exit Sub
        End If

        ''Restore form position and size:
        'If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
        'If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
        'If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
        'If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

        txtFileName.Text = FileName

        If Settings.<CoordinateConversionSettings>.<SettingsName>.Value <> Nothing Then txtDataName.Text = Settings.<CoordinateConversionSettings>.<SettingsName>.Value
        If Settings.<CoordinateConversionSettings>.<SettingsDescription>.Value <> Nothing Then txtDescription.Text = Settings.<CoordinateConversionSettings>.<SettingsDescription>.Value

        'Add code to read other saved setting here:
        If Settings.<CoordinateConversionSettings>.<SelectedTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<CoordinateConversionSettings>.<SelectedTabIndex>.Value

        Dim InputEntryCoordType As String
        If Settings.<CoordinateConversionSettings>.<InputEntryCoordType>.Value <> Nothing Then
            InputEntryCoordType = Settings.<CoordinateConversionSettings>.<InputEntryCoordType>.Value
            Select Case InputEntryCoordType
                Case "Projected"
                    rbInputEastNorth.Checked = True
                Case "Geographic"
                    rbInputLongLat.Checked = True
                Case "Cartesian"
                    rbInputXYZ.Checked = True
            End Select
        End If

        If Settings.<CoordinateConversionSettings>.<SelectedInputTabIndex>.Value <> Nothing Then TabControl2.SelectedIndex = Settings.<CoordinateConversionSettings>.<SelectedInputTabIndex>.Value
        If Settings.<CoordinateConversionSettings>.<InputCrsQuery>.Value <> Nothing Then
            txtInputCrsQuery.Text = Settings.<CoordinateConversionSettings>.<InputCrsQuery>.Value
            ApplyInputCrsQuery()
        End If

        If Settings.<CoordinateConversionSettings>.<SelectedOutputTabIndex>.Value <> Nothing Then TabControl3.SelectedIndex = Settings.<CoordinateConversionSettings>.<SelectedOutputTabIndex>.Value
        If Settings.<CoordinateConversionSettings>.<OutputCrsQuery>.Value <> Nothing Then
            txtOutputCrsQuery.Text = Settings.<CoordinateConversionSettings>.<OutputCrsQuery>.Value
            ApplyOutputCrsQuery()
        End If

        If Settings.<CoordinateConversionSettings>.<InputQueryNameContains>.Value <> Nothing Then txtFindInput.Text = Settings.<CoordinateConversionSettings>.<InputQueryNameContains>.Value
        If Settings.<CoordinateConversionSettings>.<OutputQueryNameContains>.Value <> Nothing Then txtFindOutput.Text = Settings.<CoordinateConversionSettings>.<OutputQueryNameContains>.Value

        If Settings.<CoordinateConversionSettings>.<InputCrsCode>.Value <> Nothing Then
            Conversion.InputCrs.Code = Settings.<CoordinateConversionSettings>.<InputCrsCode>.Value
            Conversion.InputCrs.GetAllSourceTargetCoordOps()
            txtInputCrsCode.Text = Conversion.InputCrs.Code
        End If
        If Settings.<CoordinateConversionSettings>.<SplitDist1>.Value <> Nothing Then
            SplitDist1 = Settings.<CoordinateConversionSettings>.<SplitDist1>.Value
            SplitContainer1.SplitterDistance = SplitDist1
        End If

        If Settings.<CoordinateConversionSettings>.<SplitDist2>.Value <> Nothing Then
            SplitDist2 = Settings.<CoordinateConversionSettings>.<SplitDist2>.Value
            SplitContainer2.SplitterDistance = SplitDist2
        End If

        If Settings.<CoordinateConversionSettings>.<OutputCrsCode>.Value <> Nothing Then
            Conversion.OutputCrs.Code = Settings.<CoordinateConversionSettings>.<OutputCrsCode>.Value
            Conversion.OutputCrs.GetAllSourceTargetCoordOps()
            txtOutputCrsCode.Text = Conversion.OutputCrs.Code
        End If

        If Settings.<CoordinateConversionSettings>.<SplitDist3>.Value <> Nothing Then
            SplitDist3 = Settings.<CoordinateConversionSettings>.<SplitDist3>.Value
            SplitContainer3.SplitterDistance = SplitDist3
        End If

        If Settings.<CoordinateConversionSettings>.<SplitDist4>.Value <> Nothing Then
            SplitDist4 = Settings.<CoordinateConversionSettings>.<SplitDist4>.Value
            SplitContainer4.SplitterDistance = SplitDist4
        End If


        If Settings.<CoordinateConversionSettings>.<SelInputRowNo>.Value <> Nothing Then
            SelInputRowNo = Settings.<CoordinateConversionSettings>.<SelInputRowNo>.Value
            If dgvInputCrsList.RowCount > SelInputRowNo And SelInputRowNo > -1 Then
                dgvInputCrsList.ClearSelection()
                dgvInputCrsList.Rows(SelInputRowNo).Selected = True
            End If
        End If
        If Settings.<CoordinateConversionSettings>.<SelOutputRowNo>.Value <> Nothing Then
            SelOutputRowNo = Settings.<CoordinateConversionSettings>.<SelOutputRowNo>.Value
            If dgvOutputCrsList.RowCount > SelOutputRowNo And SelOutputRowNo > -1 Then
                dgvOutputCrsList.ClearSelection()
                dgvOutputCrsList.Rows(SelOutputRowNo).Selected = True
            End If
        End If

        If Settings.<CoordinateConversionSettings>.<InputDegreeDisplay>.Value <> Nothing Then
            Dim InputDegreeDisplay As String = Settings.<CoordinateConversionSettings>.<InputDegreeDisplay>.Value
            If InputDegreeDisplay = "Deg-Min-Sec" Then
                rbInputDMS.Checked = True
            Else
                rbInputDecDegrees.Checked = True
            End If
        End If

        If Settings.<CoordinateConversionSettings>.<InputDegreesDecPlaces>.Value <> Nothing Then
            txtInputDegreeDecPlaces.Text = Settings.<CoordinateConversionSettings>.<InputDegreesDecPlaces>.Value
            dgvInputLocations.Columns(2).DefaultCellStyle.Format = "F" & txtInputDegreeDecPlaces.Text
        End If

        If Settings.<CoordinateConversionSettings>.<InputSecondsDecPlaces>.Value <> Nothing Then
            txtInputSecondsDecPlaces.Text = Settings.<CoordinateConversionSettings>.<InputSecondsDecPlaces>.Value
            Conversion.InputCrs.Coord.DegMinSecDecimalPlaces = txtInputSecondsDecPlaces.Text
        End If

        If Settings.<CoordinateConversionSettings>.<InputShowDmsSymbols>.Value <> Nothing Then
            chkInputDmsSymbols.Checked = Settings.<CoordinateConversionSettings>.<InputShowDmsSymbols>.Value
        End If

        If Settings.<CoordinateConversionSettings>.<InputHeightFormat>.Value <> Nothing Then
            txtInputHeightFormat.Text = Settings.<CoordinateConversionSettings>.<InputHeightFormat>.Value
            dgvInputLocations.Columns(4).DefaultCellStyle.Format = txtInputHeightFormat.Text
        End If

        If Settings.<CoordinateConversionSettings>.<InputProjectedFormat>.Value <> Nothing Then
            txtInputProjFormat.Text = Settings.<CoordinateConversionSettings>.<InputProjectedFormat>.Value
            dgvInputLocations.Columns(0).DefaultCellStyle.Format = txtInputProjFormat.Text
            dgvInputLocations.Columns(1).DefaultCellStyle.Format = txtInputProjFormat.Text
        End If

        If Settings.<CoordinateConversionSettings>.<InputCartesianFormat>.Value <> Nothing Then
            txtInputCartFormat.Text = Settings.<CoordinateConversionSettings>.<InputCartesianFormat>.Value
            dgvInputLocations.Columns(5).DefaultCellStyle.Format = txtInputCartFormat.Text
            dgvInputLocations.Columns(6).DefaultCellStyle.Format = txtInputCartFormat.Text
            dgvInputLocations.Columns(7).DefaultCellStyle.Format = txtInputCartFormat.Text
        End If

        If Settings.<CoordinateConversionSettings>.<InputLocationValues>.Value <> Nothing Then
            Dim InputLocationValues = From Item In Settings.<CoordinateConversionSettings>.<InputLocationValues>.<Row>
            For Each Item In InputLocationValues
                dgvInputLocations.Rows.Add(Val(Item.<Easting>.Value.Replace(",", "")), Val(Item.<Northing>.Value.Replace(",", "")), Item.<Longitude>.Value, Item.<Latitude>.Value, Item.<EllipsoidalHeight>.Value, Val(Item.<X>.Value.Replace(",", "")), Val(Item.<Y>.Value.Replace(",", "")), Val(Item.<Z>.Value.Replace(",", "")))
            Next
        End If
        dgvInputLocations.AutoResizeColumns()

        If Settings.<CoordinateConversionSettings>.<OutputDegreeDisplay>.Value <> Nothing Then
            Dim OutputDegreeDisplay As String = Settings.<CoordinateConversionSettings>.<OutputDegreeDisplay>.Value
            If OutputDegreeDisplay = "Deg-Min-Sec" Then
                rbOutputDMS.Checked = True
            Else
                rbOutputDecDegrees.Checked = True
            End If
        End If

        If Settings.<CoordinateConversionSettings>.<OutputDegreesDecPlaces>.Value <> Nothing Then
            txtOutputDegreeDecPlaces.Text = Settings.<CoordinateConversionSettings>.<OutputDegreesDecPlaces>.Value
            dgvOutputLocations.Columns(2).DefaultCellStyle.Format = "F" & txtOutputDegreeDecPlaces.Text
        End If

        If Settings.<CoordinateConversionSettings>.<OutputSecondsDecPlaces>.Value <> Nothing Then
            txtOutputSecondsDecPlaces.Text = Settings.<CoordinateConversionSettings>.<OutputSecondsDecPlaces>.Value
            Conversion.OutputCrs.Coord.DegMinSecDecimalPlaces = txtOutputSecondsDecPlaces.Text
        End If

        If Settings.<CoordinateConversionSettings>.<OutputShowDmsSymbols>.Value <> Nothing Then
            chkOutputDmsSymbols.Checked = Settings.<CoordinateConversionSettings>.<OutputShowDmsSymbols>.Value
        End If

        If Settings.<CoordinateConversionSettings>.<OutputHeightFormat>.Value <> Nothing Then
            txtOutputHeightFormat.Text = Settings.<CoordinateConversionSettings>.<OutputHeightFormat>.Value
            dgvOutputLocations.Columns(4).DefaultCellStyle.Format = txtOutputHeightFormat.Text
        End If

        If Settings.<CoordinateConversionSettings>.<OutputProjectedFormat>.Value <> Nothing Then
            txtOutputProjFormat.Text = Settings.<CoordinateConversionSettings>.<OutputProjectedFormat>.Value
            dgvOutputLocations.Columns(0).DefaultCellStyle.Format = txtOutputProjFormat.Text
            dgvOutputLocations.Columns(1).DefaultCellStyle.Format = txtOutputProjFormat.Text
        End If

        If Settings.<CoordinateConversionSettings>.<OutputCartesianFormat>.Value <> Nothing Then
            txtOutputCartFormat.Text = Settings.<CoordinateConversionSettings>.<OutputCartesianFormat>.Value
            dgvOutputLocations.Columns(5).DefaultCellStyle.Format = txtOutputCartFormat.Text
            dgvOutputLocations.Columns(6).DefaultCellStyle.Format = txtOutputCartFormat.Text
            dgvOutputLocations.Columns(7).DefaultCellStyle.Format = txtOutputCartFormat.Text
        End If

        If Settings.<CoordinateConversionSettings>.<OutputLocationValues>.Value <> Nothing Then
            Dim OutputLocationValues = From Item In Settings.<CoordinateConversionSettings>.<OutputLocationValues>.<Row>
            For Each Item In OutputLocationValues
                dgvOutputLocations.Rows.Add(Val(Item.<Easting>.Value.Replace(",", "")), Val(Item.<Northing>.Value.Replace(",", "")), Item.<Longitude>.Value, Item.<Latitude>.Value, Item.<EllipsoidalHeight>.Value, Val(Item.<X>.Value.Replace(",", "")), Val(Item.<Y>.Value.Replace(",", "")), Val(Item.<Z>.Value.Replace(",", "")))
            Next
        End If
        dgvInputLocations.AutoResizeColumns()

        If Settings.<CoordinateConversionSettings>.<SelectedDatumTransTabIndex>.Value <> Nothing Then TabControl4.SelectedIndex = Settings.<CoordinateConversionSettings>.<SelectedDatumTransTabIndex>.Value
        If Settings.<CoordinateConversionSettings>.<SelectedViaWgs84TabIndex>.Value <> Nothing Then TabControl5.SelectedIndex = Settings.<CoordinateConversionSettings>.<SelectedViaWgs84TabIndex>.Value
        If Settings.<CoordinateConversionSettings>.<SplitDist5>.Value <> Nothing Then
            SplitDist5 = Settings.<CoordinateConversionSettings>.<SplitDist5>.Value
            SplitContainer5.SplitterDistance = SplitDist5
        End If
        If Settings.<CoordinateConversionSettings>.<SplitDist6>.Value <> Nothing Then
            SplitDist6 = Settings.<CoordinateConversionSettings>.<SplitDist6>.Value
            SplitContainer6.SplitterDistance = SplitDist6
        End If

        If Settings.<CoordinateConversionSettings>.<SplitDist7>.Value <> Nothing Then
            SplitDist7 = Settings.<CoordinateConversionSettings>.<SplitDist7>.Value
            SplitContainer7.SplitterDistance = SplitDist7
        End If

        If Settings.<CoordinateConversionSettings>.<SplitDist8>.Value <> Nothing Then
            SplitDist8 = Settings.<CoordinateConversionSettings>.<SplitDist8>.Value
            SplitContainer8.SplitterDistance = SplitDist8
        End If
        If Settings.<CoordinateConversionSettings>.<SplitDist9>.Value <> Nothing Then
            SplitDist9 = Settings.<CoordinateConversionSettings>.<SplitDist9>.Value
            SplitContainer9.SplitterDistance = SplitDist9
        End If
        If Settings.<CoordinateConversionSettings>.<SplitDist10>.Value <> Nothing Then
            SplitDist10 = Settings.<CoordinateConversionSettings>.<SplitDist10>.Value
            SplitContainer10.SplitterDistance = SplitDist10
        End If
        If Settings.<CoordinateConversionSettings>.<SplitDist11>.Value <> Nothing Then
            SplitDist11 = Settings.<CoordinateConversionSettings>.<SplitDist11>.Value
            SplitContainer11.SplitterDistance = SplitDist11
        End If

        If Settings.<CoordinateConversionSettings>.<SplitDist12>.Value <> Nothing Then
            SplitDist12 = Settings.<CoordinateConversionSettings>.<SplitDist12>.Value
            SplitContainer12.SplitterDistance = SplitDist12
        End If

        If Settings.<CoordinateConversionSettings>.<SplitDist13>.Value <> Nothing Then
            SplitDist13 = Settings.<CoordinateConversionSettings>.<SplitDist13>.Value
            SplitContainer13.SplitterDistance = SplitDist13
        End If

        If Settings.<CoordinateConversionSettings>.<SplitDist14>.Value <> Nothing Then
            SplitDist14 = Settings.<CoordinateConversionSettings>.<SplitDist14>.Value
            SplitContainer14.SplitterDistance = SplitDist14
        End If
        If Settings.<CoordinateConversionSettings>.<SplitDist15>.Value <> Nothing Then
            SplitDist15 = Settings.<CoordinateConversionSettings>.<SplitDist15>.Value
            SplitContainer15.SplitterDistance = SplitDist15
        End If
        If Settings.<CoordinateConversionSettings>.<DirectCoordOpCode>.Value <> Nothing Then DirectCoordOpCode = Settings.<CoordinateConversionSettings>.<DirectCoordOpCode>.Value
        If Settings.<CoordinateConversionSettings>.<InputToWgs84CoordOpCode>.Value <> Nothing Then InputToWgs84CoordOpCode = Settings.<CoordinateConversionSettings>.<InputToWgs84CoordOpCode>.Value
        If Settings.<CoordinateConversionSettings>.<Wgs84ToOutputCoordOpCode>.Value <> Nothing Then Wgs84ToOutputCoordOpCode = Settings.<CoordinateConversionSettings>.<Wgs84ToOutputCoordOpCode>.Value
        Dim EntryCoordType As String
        If Settings.<CoordinateConversionSettings>.<EntryCoordType>.Value <> Nothing Then
            EntryCoordType = Settings.<CoordinateConversionSettings>.<EntryCoordType>.Value
            'Select Case Settings.<FormSettings>.<EntryCoordType>.Value
            Select Case EntryCoordType
                Case "Projected"
                    rbEnterInputEastNorth.Checked = True
                Case "Geographic"
                    rbEnterInputLongLat.Checked = True
                Case "Cartesian"
                    rbEnterInputXYZ.Checked = True
            End Select
        End If
        If Settings.<CoordinateConversionSettings>.<ShowPointNumber>.Value <> Nothing Then chkShowPointNumber.Checked = Settings.<CoordinateConversionSettings>.<ShowPointNumber>.Value
        If Settings.<CoordinateConversionSettings>.<ShowPointName>.Value <> Nothing Then chkShowPointName.Checked = Settings.<CoordinateConversionSettings>.<ShowPointName>.Value
        If Settings.<CoordinateConversionSettings>.<ShowPointDescription>.Value <> Nothing Then chkShowPointDescription.Checked = Settings.<CoordinateConversionSettings>.<ShowPointDescription>.Value
        If Settings.<CoordinateConversionSettings>.<ShowInputEastingNorthing>.Value <> Nothing Then chkShowInputEastNorth.Checked = Settings.<CoordinateConversionSettings>.<ShowInputEastingNorthing>.Value
        If Settings.<CoordinateConversionSettings>.<ShowInputLongitudeLatitude>.Value <> Nothing Then chkShowInputLongLat.Checked = Settings.<CoordinateConversionSettings>.<ShowInputLongitudeLatitude>.Value
        If Settings.<CoordinateConversionSettings>.<ShowInputXYS>.Value <> Nothing Then chkShowInputXYZ.Checked = Settings.<CoordinateConversionSettings>.<ShowInputXYS>.Value
        If Settings.<CoordinateConversionSettings>.<ShowWgs84XYZ>.Value <> Nothing Then chkShowWgs84XYZ.Checked = Settings.<CoordinateConversionSettings>.<ShowWgs84XYZ>.Value
        If Settings.<CoordinateConversionSettings>.<ShowOutputEastingNorthing>.Value <> Nothing Then chkShowOutputEastNorth.Checked = Settings.<CoordinateConversionSettings>.<ShowOutputEastingNorthing>.Value
        If Settings.<CoordinateConversionSettings>.<ShowOutputLongitudeLatitude>.Value <> Nothing Then chkShowOutputLongLat.Checked = Settings.<CoordinateConversionSettings>.<ShowOutputLongitudeLatitude>.Value
        If Settings.<CoordinateConversionSettings>.<ShowOutputXYZ>.Value <> Nothing Then chkShowOutputXYZ.Checked = Settings.<CoordinateConversionSettings>.<ShowOutputXYZ>.Value
        If Settings.<CoordinateConversionSettings>.<ProjectedFormat>.Value <> Nothing Then txtProjFormat.Text = Settings.<CoordinateConversionSettings>.<ProjectedFormat>.Value
        If Settings.<CoordinateConversionSettings>.<CartesianFormat>.Value <> Nothing Then txtCartFormat.Text = Settings.<CoordinateConversionSettings>.<CartesianFormat>.Value
        If Settings.<CoordinateConversionSettings>.<DecDegreesDecPlaces>.Value <> Nothing Then
            txtDegreeDecPlaces.Text = Settings.<CoordinateConversionSettings>.<DecDegreesDecPlaces>.Value
            If rbDecDegrees.Checked Then
                Try
                    txtDegreeDecPlaces.Text = Int(txtDegreeDecPlaces.Text.Trim)
                    Dim Format As String = "F" & txtDegreeDecPlaces.Text.Trim
                    dgvConversion.Columns(5).DefaultCellStyle.Format = Format
                    dgvConversion.Columns(6).DefaultCellStyle.Format = Format
                    dgvConversion.Columns(17).DefaultCellStyle.Format = Format
                    dgvConversion.Columns(18).DefaultCellStyle.Format = Format
                Catch ex As Exception

                End Try
            End If
        End If
        If Settings.<CoordinateConversionSettings>.<DmsSecondsDecPlaces>.Value <> Nothing Then
            txtSecondsDecPlaces.Text = Settings.<CoordinateConversionSettings>.<DmsSecondsDecPlaces>.Value
            Conversion.InputCrs.Coord.DegMinSecDecimalPlaces = txtSecondsDecPlaces.Text
            Conversion.OutputCrs.Coord.DegMinSecDecimalPlaces = txtSecondsDecPlaces.Text
        End If
        If Settings.<CoordinateConversionSettings>.<HeightFormat>.Value <> Nothing Then
            txtHeightFormat.Text = Settings.<CoordinateConversionSettings>.<HeightFormat>.Value
        End If

        If Settings.<CoordinateConversionSettings>.<ShowDegMinSec>.Value <> Nothing Then rbDMS.Checked = Settings.<CoordinateConversionSettings>.<ShowDegMinSec>.Value
        If Settings.<CoordinateConversionSettings>.<ShowDegMinSecSymbols>.Value <> Nothing Then chkDmsSymbols.Checked = Settings.<CoordinateConversionSettings>.<ShowDegMinSecSymbols>.Value
        If EntryCoordType = "Projected" Then
            Dim RowNo As Integer
            If Settings.<CoordinateConversionSettings>.<DatumTransInputData>.Value <> Nothing Then
                Dim InputPoints = From Item In Settings.<CoordinateConversionSettings>.<DatumTransInputData>.<Row>
                For Each Item In InputPoints
                    RowNo = dgvConversion.Rows.Add()
                    If Item.<PointNumber>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(0).Value = Item.<PointNumber>.Value
                    If Item.<PointName>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(1).Value = Item.<PointName>.Value
                    If Item.<PointDescription>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(2).Value = Item.<PointDescription>.Value
                    dgvConversion.Rows(RowNo).Cells(3).Value = Item.<Easting>.Value
                    dgvConversion.Rows(RowNo).Cells(4).Value = Item.<Northing>.Value
                Next
            End If

        ElseIf EntryCoordType = "Geographic" Then

            Dim RowNo As Integer

            If Settings.<CoordinateConversionSettings>.<DatumTransInputData>.Value <> Nothing Then
                Dim InputPoints = From Item In Settings.<CoordinateConversionSettings>.<DatumTransInputData>.<Row>
                For Each Item In InputPoints
                    RowNo = dgvConversion.Rows.Add()
                    If Item.<PointNumber>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(0).Value = Item.<PointNumber>.Value
                    If Item.<PointName>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(1).Value = Item.<PointName>.Value
                    If Item.<PointDescription>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(2).Value = Item.<PointDescription>.Value
                    dgvConversion.Rows(RowNo).Cells(5).Value = Item.<Longitude>.Value
                    dgvConversion.Rows(RowNo).Cells(6).Value = Item.<Latitude>.Value
                    dgvConversion.Rows(RowNo).Cells(7).Value = Item.<Height>.Value
                Next
            End If

        ElseIf EntryCoordType = "Cartesian" Then
            Dim RowNo As Integer
            If Settings.<CoordinateConversionSettings>.<DatumTransInputData>.Value <> Nothing Then
                Dim InputPoints = From Item In Settings.<CoordinateConversionSettings>.<DatumTransInputData>.<Row>
                For Each Item In InputPoints
                    RowNo = dgvConversion.Rows.Add()
                    If Item.<PointNumber>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(0).Value = Item.<PointNumber>.Value
                    If Item.<PointName>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(1).Value = Item.<PointName>.Value
                    If Item.<PointDescription>.Value <> Nothing Then dgvConversion.Rows(RowNo).Cells(2).Value = Item.<PointDescription>.Value
                    dgvConversion.Rows(RowNo).Cells(8).Value = Item.<X>.Value
                    dgvConversion.Rows(RowNo).Cells(9).Value = Item.<Y>.Value
                    dgvConversion.Rows(RowNo).Cells(10).Value = Item.<Z>.Value
                Next
            End If

        End If
        ApplyDatumTransFormats()

        UpdateDatumTransTable()

        ShowInputCrsInfo()
        ShowOutputCrsInfo()
        DisplayDirectTransformationOptions()
        DisplayInputToWgs84TransOptions()
        DisplayWgs84ToOutputTransOptions()

        'DisplayDirectTransformationOptions()
        DisplayTransformationOptions(Conversion.InputCrs, Conversion.OutputCrs)

        'Display the selected Datum Transformation methods:
        SelectDirectTransOpCode(DirectCoordOpCode)
        SelectInputToWgs84TransOpCode(InputToWgs84CoordOpCode)
        SelectWgs84ToOutputTransOpCode(Wgs84ToOutputCoordOpCode)

        If txtInputCrsQuery.Text.Trim = "" Then txtInputCrsQuery.Text = "Select COORD_REF_SYS_CODE, COORD_REF_SYS_NAME, COORD_REF_SYS_KIND, REMARKS From [Coordinate Reference System]"
        If txtOutputCrsQuery.Text.Trim = "" Then txtOutputCrsQuery.Text = "Select COORD_REF_SYS_CODE, COORD_REF_SYS_NAME, COORD_REF_SYS_KIND, REMARKS From [Coordinate Reference System]"

        ReCalcDatumTransTable()

        FindDefaultDatumTrans()

        dgvConversion.AutoResizeColumns()
        dgvInputLocations.AutoResizeColumns()
        dgvOutputLocations.AutoResizeColumns()

        Modified = False

    End Sub


    Private Sub dgvInputLocations_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvInputLocations.CellEndEdit
        'The cell has been edited - update the coordinates.

        Dim ColHeader As String = dgvInputLocations.Columns(e.ColumnIndex).HeaderText
        Select Case ColHeader
            Case "Easting"
                If rbInputEastNorth.Checked Then 'This is a valid entry coordinate.
                    'Check that the Northing value is valid:
                    'If dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).ToString.Trim = "" Then  'There is no Northing value
                    If dgvInputLocations.Rows(e.RowIndex).Cells(1).ToString.Trim = "" Then  'There is no Northing value

                    Else
                        'Conversion.InputCrs.Coord.SetEasting(dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.None) 'Update the new Input Easting value
                        Conversion.InputCrs.Coord.SetEasting(dgvInputLocations.Rows(e.RowIndex).Cells(0).Value, Coordinate.UpdateMode.None) 'Read the Input Easting value
                        'Conversion.InputCrs.Coord.SetNorthing(dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value, Coordinate.UpdateMode.All) 'Read the Input Northing Update all the Input coordinate types
                        Conversion.InputCrs.Coord.SetNorthing(dgvInputLocations.Rows(e.RowIndex).Cells(1).Value, Coordinate.UpdateMode.All) 'Read the Input Northing Update all the Input coordinate types
                        If rbInputDecDegrees.Checked Then
                            dgvInputLocations.Rows(e.RowIndex).Cells(2).Value = Conversion.InputCrs.Coord.Longitude
                            dgvInputLocations.Rows(e.RowIndex).Cells(3).Value = Conversion.InputCrs.Coord.Latitude
                        Else
                            dgvInputLocations.Rows(e.RowIndex).Cells(2).Value = Conversion.InputCrs.Coord.LongitudeDMS
                            dgvInputLocations.Rows(e.RowIndex).Cells(3).Value = Conversion.InputCrs.Coord.LatitudeDMS
                        End If
                        dgvInputLocations.Rows(e.RowIndex).Cells(4).Value = Conversion.InputCrs.Coord.EllipsoidalHeight
                        dgvInputLocations.Rows(e.RowIndex).Cells(5).Value = Conversion.InputCrs.Coord.X
                        dgvInputLocations.Rows(e.RowIndex).Cells(6).Value = Conversion.InputCrs.Coord.Y
                        dgvInputLocations.Rows(e.RowIndex).Cells(7).Value = Conversion.InputCrs.Coord.Z
                        dgvInputLocations.AutoResizeColumns()
                    End If
                End If
            Case "Northing"
                If rbInputEastNorth.Checked Then 'This is a valid entry coordinate.
                    'Check that the Easting value is valid:
                    'If dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).ToString.Trim = "" Then  'There is no Easting value
                    If dgvInputLocations.Rows(e.RowIndex).Cells(0).ToString.Trim = "" Then  'There is no Easting value

                    Else
                        'Conversion.InputCrs.Coord.SetEasting(dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, Coordinate.UpdateMode.None) 'Read the Input Easting value
                        Conversion.InputCrs.Coord.SetEasting(dgvInputLocations.Rows(e.RowIndex).Cells(0).Value, Coordinate.UpdateMode.None) 'Read the Input Easting value
                        'Conversion.InputCrs.Coord.SetNorthing(dgvInputLocations.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value, Coordinate.UpdateMode.All) 'Read the Input Northing Update all the Input coordinate types
                        Conversion.InputCrs.Coord.SetNorthing(dgvInputLocations.Rows(e.RowIndex).Cells(1).Value, Coordinate.UpdateMode.All) 'Read the Input Northing and Update all the Input coordinate types
                        If rbInputDecDegrees.Checked Then
                            dgvInputLocations.Rows(e.RowIndex).Cells(2).Value = Conversion.InputCrs.Coord.Longitude
                            dgvInputLocations.Rows(e.RowIndex).Cells(3).Value = Conversion.InputCrs.Coord.Latitude
                        Else
                            dgvInputLocations.Rows(e.RowIndex).Cells(2).Value = Conversion.InputCrs.Coord.LongitudeDMS
                            dgvInputLocations.Rows(e.RowIndex).Cells(3).Value = Conversion.InputCrs.Coord.LatitudeDMS
                        End If
                        dgvInputLocations.Rows(e.RowIndex).Cells(4).Value = Conversion.InputCrs.Coord.EllipsoidalHeight
                        dgvInputLocations.Rows(e.RowIndex).Cells(5).Value = Conversion.InputCrs.Coord.X
                        dgvInputLocations.Rows(e.RowIndex).Cells(6).Value = Conversion.InputCrs.Coord.Y
                        dgvInputLocations.Rows(e.RowIndex).Cells(7).Value = Conversion.InputCrs.Coord.Z
                        dgvInputLocations.AutoResizeColumns()
                    End If
                End If

            Case "Longitude"
                If rbInputLongLat.Checked Then 'This is a valid entry coordinate.
                    'Check that the Ellipsoidal Height value is valid:
                    If dgvInputLocations.Rows(e.RowIndex).Cells(4).ToString.Trim = "" Then  'There is no Ellipsoidal Height value.
                        dgvInputLocations.Rows(e.RowIndex).Cells(4).Value = 0 'Set the Ellipsoidal Height to the default value of 0.
                    End If
                    'Check that the Latitude value is valid:
                    If dgvInputLocations.Rows(e.RowIndex).Cells(3).ToString.Trim = "" Then  'There is no Latitude value

                    Else
                        Conversion.InputCrs.Coord.SetLatitude(dgvInputLocations.Rows(e.RowIndex).Cells(3).Value, Coordinate.UpdateMode.None) 'Read the Input Latitude value
                        Conversion.InputCrs.Coord.SetEllipsoidalHeight(dgvInputLocations.Rows(e.RowIndex).Cells(4).Value, Coordinate.UpdateMode.None) 'Read the Input Ellipsoidal Height value
                        Conversion.InputCrs.Coord.SetLongitude(dgvInputLocations.Rows(e.RowIndex).Cells(2).Value, Coordinate.UpdateMode.All) 'Read the Input Longitude value and Update all the Input coordinate types
                        dgvInputLocations.Rows(e.RowIndex).Cells(0).Value = Conversion.InputCrs.Coord.Easting
                        dgvInputLocations.Rows(e.RowIndex).Cells(1).Value = Conversion.InputCrs.Coord.Northing
                        dgvInputLocations.Rows(e.RowIndex).Cells(5).Value = Conversion.InputCrs.Coord.X
                        dgvInputLocations.Rows(e.RowIndex).Cells(6).Value = Conversion.InputCrs.Coord.Y
                        dgvInputLocations.Rows(e.RowIndex).Cells(7).Value = Conversion.InputCrs.Coord.Z
                        dgvInputLocations.AutoResizeColumns()
                    End If
                End If

            Case "Latitude"
                If rbInputLongLat.Checked Then 'This is a valid entry coordinate.
                    'Check that the Ellipsoidal Height value is valid:
                    If dgvInputLocations.Rows(e.RowIndex).Cells(4).ToString.Trim = "" Then  'There is no Ellipsoidal Height value.
                        dgvInputLocations.Rows(e.RowIndex).Cells(4).Value = 0 'Set the Ellipsoidal Height to the default value of 0.
                    End If
                    'Check that the Longitude value is valid:
                    If dgvInputLocations.Rows(e.RowIndex).Cells(2).ToString.Trim = "" Then  'There is no Longitude value

                    Else
                        Conversion.InputCrs.Coord.SetLatitude(dgvInputLocations.Rows(e.RowIndex).Cells(3).Value, Coordinate.UpdateMode.None) 'Read the Input Latitude value
                        Conversion.InputCrs.Coord.SetEllipsoidalHeight(dgvInputLocations.Rows(e.RowIndex).Cells(4).Value, Coordinate.UpdateMode.None) 'Read the Input Ellipsoidal Height value
                        Conversion.InputCrs.Coord.SetLongitude(dgvInputLocations.Rows(e.RowIndex).Cells(2).Value, Coordinate.UpdateMode.All) 'Read the Input Longitude value and Update all the Input coordinate types
                        dgvInputLocations.Rows(e.RowIndex).Cells(0).Value = Conversion.InputCrs.Coord.Easting
                        dgvInputLocations.Rows(e.RowIndex).Cells(1).Value = Conversion.InputCrs.Coord.Northing
                        dgvInputLocations.Rows(e.RowIndex).Cells(5).Value = Conversion.InputCrs.Coord.X
                        dgvInputLocations.Rows(e.RowIndex).Cells(6).Value = Conversion.InputCrs.Coord.Y
                        dgvInputLocations.Rows(e.RowIndex).Cells(7).Value = Conversion.InputCrs.Coord.Z
                        dgvInputLocations.AutoResizeColumns()
                    End If
                End If

            Case "Ellipsoidal Height"
                If rbInputLongLat.Checked Then 'This is a valid entry coordinate.
                    'Check that the Longitude and Latitude values are valid:
                    If dgvInputLocations.Rows(e.RowIndex).Cells(2).ToString.Trim = "" And dgvInputLocations.Rows(e.RowIndex).Cells(3).ToString.Trim = "" Then  'There is no Longitude or Latitude value

                    Else
                        Conversion.InputCrs.Coord.SetLatitude(dgvInputLocations.Rows(e.RowIndex).Cells(3).Value, Coordinate.UpdateMode.None) 'Read the Input Latitude value
                        Conversion.InputCrs.Coord.SetEllipsoidalHeight(dgvInputLocations.Rows(e.RowIndex).Cells(4).Value, Coordinate.UpdateMode.None) 'Read the Input Ellipsoidal Height value
                        Conversion.InputCrs.Coord.SetLongitude(dgvInputLocations.Rows(e.RowIndex).Cells(2).Value, Coordinate.UpdateMode.All) 'Read the Input Longitude value and Update all the Input coordinate types
                        dgvInputLocations.Rows(e.RowIndex).Cells(0).Value = Conversion.InputCrs.Coord.Easting
                        dgvInputLocations.Rows(e.RowIndex).Cells(1).Value = Conversion.InputCrs.Coord.Northing
                        dgvInputLocations.Rows(e.RowIndex).Cells(5).Value = Conversion.InputCrs.Coord.X
                        dgvInputLocations.Rows(e.RowIndex).Cells(6).Value = Conversion.InputCrs.Coord.Y
                        dgvInputLocations.Rows(e.RowIndex).Cells(7).Value = Conversion.InputCrs.Coord.Z
                        dgvInputLocations.AutoResizeColumns()
                    End If
                End If

            Case "X"

            Case "Y"

            Case "Z"

            Case Else

        End Select

    End Sub



    Private Sub SelectDirectTransOpCode(CodeNo As Integer)
        'Select the Direct Datum Transformation Operation
        Dim ApplyReverse As Boolean
        For Each Row As DataGridViewRow In dgvDirectDTOps.Rows
            If Row.Cells(2).Value = CodeNo Then
                Row.Selected = True
                ApplyReverse = Row.Cells(12).Value
                Conversion.DatumTrans.GetDirectDatumTransCoordOp(CodeNo)
                Conversion.DatumTrans.DirectMethodApplyReverse = ApplyReverse
                DisplayDirectDatumTransMethod()
                Exit For
            End If
        Next

    End Sub

    Private Sub SelectInputToWgs84TransOpCode(CodeNo As Integer)
        'Select the Input to WGS 84  Datum Transformation Operation
        Dim ApplyReverse As Boolean
        For Each Row As DataGridViewRow In dgvInputToWgs84DTOps.Rows
            If Row.Cells(2).Value = CodeNo Then
                Row.Selected = True
                ApplyReverse = Row.Cells(12).Value
                Conversion.DatumTrans.GetInputToWgs84DatumTransCoordOp(CodeNo)
                Conversion.DatumTrans.InputToWgs84MethodApplyReverse = ApplyReverse
                DisplayInputToWgs84DatumTransMethod()
                Exit For
            End If
        Next

    End Sub

    Private Sub SelectWgs84ToOutputTransOpCode(CodeNo As Integer)
        'Select the WGS 84 to Output Datum Transformation Operation
        Dim ApplyReverse As Boolean
        For Each Row As DataGridViewRow In dgvWgs84ToOutputDTOps.Rows
            If Row.Cells(2).Value = CodeNo Then
                Row.Selected = True
                ApplyReverse = Row.Cells(12).Value
                Conversion.DatumTrans.GetWgs84ToOutputDatumTransCoordOp(CodeNo)
                Conversion.DatumTrans.Wgs84ToOutputMethodApplyReverse = ApplyReverse
                DisplayWgs84ToOutputDatumTransMethod()
                Exit For
            End If
        Next
    End Sub


    Private Sub UpdateDatumTransTable()

        If chkShowPointNumber.Checked Then dgvConversion.Columns(0).Visible = True Else dgvConversion.Columns(0).Visible = False
        If chkShowPointName.Checked Then dgvConversion.Columns(1).Visible = True Else dgvConversion.Columns(1).Visible = False
        If chkShowPointDescription.Checked Then dgvConversion.Columns(2).Visible = True Else dgvConversion.Columns(2).Visible = False
        If chkShowInputEastNorth.Checked Then
            dgvConversion.Columns(3).Visible = True
            dgvConversion.Columns(4).Visible = True
        Else
            dgvConversion.Columns(3).Visible = False
            dgvConversion.Columns(4).Visible = False
        End If
        If chkShowInputLongLat.Checked Then
            dgvConversion.Columns(5).Visible = True
            dgvConversion.Columns(6).Visible = True
            dgvConversion.Columns(7).Visible = True
        Else
            dgvConversion.Columns(5).Visible = False
            dgvConversion.Columns(6).Visible = False
            dgvConversion.Columns(7).Visible = False
        End If
        If chkShowInputXYZ.Checked Then
            dgvConversion.Columns(8).Visible = True
            dgvConversion.Columns(9).Visible = True
            dgvConversion.Columns(10).Visible = True
        Else
            dgvConversion.Columns(8).Visible = False
            dgvConversion.Columns(9).Visible = False
            dgvConversion.Columns(10).Visible = False
        End If
        If chkShowWgs84XYZ.Checked Then
            dgvConversion.Columns(11).Visible = True
            dgvConversion.Columns(12).Visible = True
            dgvConversion.Columns(13).Visible = True
        Else
            dgvConversion.Columns(11).Visible = False
            dgvConversion.Columns(12).Visible = False
            dgvConversion.Columns(13).Visible = False
        End If
        If chkShowOutputXYZ.Checked Then
            dgvConversion.Columns(14).Visible = True
            dgvConversion.Columns(15).Visible = True
            dgvConversion.Columns(16).Visible = True
        Else
            dgvConversion.Columns(14).Visible = False
            dgvConversion.Columns(15).Visible = False
            dgvConversion.Columns(16).Visible = False
        End If
        If chkShowOutputLongLat.Checked Then
            dgvConversion.Columns(17).Visible = True
            dgvConversion.Columns(18).Visible = True
            dgvConversion.Columns(19).Visible = True
        Else
            dgvConversion.Columns(17).Visible = False
            dgvConversion.Columns(18).Visible = False
            dgvConversion.Columns(19).Visible = False
        End If
        If chkShowOutputEastNorth.Checked Then
            dgvConversion.Columns(20).Visible = True
            dgvConversion.Columns(21).Visible = True
        Else
            dgvConversion.Columns(20).Visible = False
            dgvConversion.Columns(21).Visible = False
        End If
    End Sub

    Private Sub UpdateDatumTransTableInput()
        'Update the Datum Transformation Table Input columns: dgvConversion
        'This method sets the data entry columns to Read-Write with a white background.

        If rbEnterInputEastNorth.Checked Then
            For Each Col As DataGridViewColumn In dgvConversion.Columns
                'If Col.HeaderText = "Input Easting" Or Col.HeaderText = "Input Northing" Then
                If Col.HeaderText = "Input Easting" Or Col.HeaderText = "Input Northing" Or Col.HeaderText = "Point Number" Or Col.HeaderText = "Point Name" Or Col.HeaderText = "Point Description" Then
                    Col.ReadOnly = False
                    Col.DefaultCellStyle.BackColor = Color.White
                Else
                    Col.ReadOnly = True
                    Col.DefaultCellStyle.BackColor = Color.WhiteSmoke
                End If
            Next

        ElseIf rbEnterInputLongLat.Checked Then
            For Each Col As DataGridViewColumn In dgvConversion.Columns
                If Col.HeaderText = "Input Longitude" Or Col.HeaderText = "Input Latitude" Or Col.HeaderText = "Input Ellipsoidal Height" Or Col.HeaderText = "Point Number" Or Col.HeaderText = "Point Name" Or Col.HeaderText = "Point Description" Then
                    Col.ReadOnly = False
                    Col.DefaultCellStyle.BackColor = Color.White
                Else
                    Col.ReadOnly = True
                    Col.DefaultCellStyle.BackColor = Color.WhiteSmoke
                End If
            Next
        ElseIf rbEnterInputXYZ.Checked Then
            For Each Col As DataGridViewColumn In dgvConversion.Columns
                If Col.HeaderText = "Input X" Or Col.HeaderText = "Input Y" Or Col.HeaderText = "Input Z" Or Col.HeaderText = "Point Number" Or Col.HeaderText = "Point Name" Or Col.HeaderText = "Point Description" Then
                    Col.ReadOnly = False
                    Col.DefaultCellStyle.BackColor = Color.White
                Else
                    Col.ReadOnly = True
                    Col.DefaultCellStyle.BackColor = Color.WhiteSmoke
                End If
            Next
        End If
    End Sub



    Private Sub UpdateDatumTransTable(RowNo As Integer)
        'Update the calculated values in the Datum Transformation table dgvConversion
        'The row specified by RowNo is updated.

        For Each Col As DataGridViewColumn In dgvConversion.Columns
            If Col.ReadOnly = True Then
                Select Case Col.HeaderText
                    Case "Input Easting"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.InputCrs.Coord.Easting
                    Case "Input Northing"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.InputCrs.Coord.Northing
                    Case "Input Longitude"
                        If rbDMS.Checked Then
                            dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.InputCrs.Coord.LongitudeDMS
                        Else
                            dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.InputCrs.Coord.Longitude
                        End If
                    Case "Input Latitude"
                        If rbDMS.Checked Then
                            dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.InputCrs.Coord.LatitudeDMS
                        Else
                            dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.InputCrs.Coord.Latitude
                        End If

                    Case "Input Ellipsoidal Height"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.InputCrs.Coord.EllipsoidalHeight
                    Case "Input X"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.InputCrs.Coord.X
                    Case "Input Y"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.InputCrs.Coord.Y
                    Case "Input Z"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.InputCrs.Coord.Z
                    Case "WGS 84 X"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.DatumTrans.Wgs84Coord.X
                    Case "WGS 84 Y"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.DatumTrans.Wgs84Coord.Y
                    Case "WGS 84 Z"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.DatumTrans.Wgs84Coord.Z
                    Case "Output X"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.OutputCrs.Coord.X
                    Case "Output Y"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.OutputCrs.Coord.Y
                    Case "Output Z"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.OutputCrs.Coord.Z
                    Case "Output Longitude"
                        If rbDMS.Checked Then
                            dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.OutputCrs.Coord.LongitudeDMS
                        Else
                            dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.OutputCrs.Coord.Longitude
                        End If
                    Case "Output Latitude"
                        If rbDMS.Checked Then
                            dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.OutputCrs.Coord.LatitudeDMS
                        Else
                            dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.OutputCrs.Coord.Latitude
                        End If
                    Case "Output Ellipsoidal Height"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.OutputCrs.Coord.EllipsoidalHeight
                    Case "Output Easting"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.OutputCrs.Coord.Easting
                    Case "Output Northing"
                        dgvConversion.Rows(RowNo).Cells(Col.Index).Value = Conversion.OutputCrs.Coord.Northing
                End Select
            End If
        Next

        dgvConversion.AutoResizeColumns()

    End Sub

    Private Sub ReCalcDatumTransTable()
        'Recalculate location values in the Datum Transformation table.

        dgvConversion.AllowUserToAddRows = False

        If rbEnterInputEastNorth.Checked Then
            'Dim EastingCol As Integer
            'Dim NorthingCol As Integer
            'Dim RowNo As Integer
            'For Each Col As DataGridViewColumn In dgvConversion.Columns
            '    If Col.HeaderText = "Input Easting" Then EastingCol = Col.Index
            '    If Col.HeaderText = "Input Northing" Then NorthingCol = Col.Index
            'Next
            For Each Row As DataGridViewRow In dgvConversion.Rows
                Conversion.InputCrs.Coord.SetEasting(Row.Cells(3).Value, Coordinate.UpdateMode.None) 'Update the new Input Easting value
                Conversion.InputCrs.Coord.SetNorthing(Row.Cells(4).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Northing value
                UpdateDatumTransTable(Row.Index)
            Next

        ElseIf rbEnterInputLongLat.Checked Then
            'Dim LongitudeCol As Integer
            'Dim LatitudeCol As Integer
            'Dim HeightCol As Integer
            'Dim RowNo As Integer
            'For Each Col As DataGridViewColumn In dgvConversion.Columns
            '    If Col.HeaderText = "Input Longitude" Then LongitudeCol = Col.Index
            '    If Col.HeaderText = "Input Latitude" Then LatitudeCol = Col.Index
            '    If Col.HeaderText = "Input Ellipsoidal Height" Then HeightCol = Col.Index
            'Next
            For Each Row As DataGridViewRow In dgvConversion.Rows
                Conversion.InputCrs.Coord.SetLongitude(Row.Cells(5).Value, Coordinate.UpdateMode.None) 'Update the new Input Longitude value
                Conversion.InputCrs.Coord.SetLatitude(Row.Cells(6).Value, Coordinate.UpdateMode.None) 'Update the new Input Latitude value
                Conversion.InputCrs.Coord.SetEllipsoidalHeight(Row.Cells(7).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Ellipsoidal Height value
                UpdateDatumTransTable(Row.Index)
            Next


        ElseIf rbEnterInputXYZ.Checked Then
            'Dim XCol As Integer
            'Dim YCol As Integer
            'Dim ZCol As Integer
            'Dim RowNo As Integer
            'For Each Col As DataGridViewColumn In dgvConversion.Columns
            '    If Col.HeaderText = "Input X" Then XCol = Col.Index
            '    If Col.HeaderText = "Input Y" Then YCol = Col.Index
            '    If Col.HeaderText = "Input Z" Then ZCol = Col.Index
            'Next
            For Each Row As DataGridViewRow In dgvConversion.Rows
                Conversion.InputCrs.Coord.SetLongitude(Row.Cells(8).Value, Coordinate.UpdateMode.None) 'Update the new Input X value
                Conversion.InputCrs.Coord.SetLatitude(Row.Cells(9).Value, Coordinate.UpdateMode.None) 'Update the new Input Y value
                Conversion.InputCrs.Coord.SetEllipsoidalHeight(Row.Cells(10).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Z value
                UpdateDatumTransTable(Row.Index)
            Next

        End If

        dgvConversion.AllowUserToAddRows = True

    End Sub



    Private Sub ReCalcDatumTransTable_OLD()
        'Recalculate location values in the Datum Transformation table.

        dgvConversion.AllowUserToAddRows = False

        If rbEnterInputEastNorth.Checked Then
            Dim EastingCol As Integer
            Dim NorthingCol As Integer
            Dim RowNo As Integer
            For Each Col As DataGridViewColumn In dgvConversion.Columns
                If Col.HeaderText = "Input Easting" Then EastingCol = Col.Index
                If Col.HeaderText = "Input Northing" Then NorthingCol = Col.Index
            Next
            For Each Row As DataGridViewRow In dgvConversion.Rows
                Conversion.InputCrs.Coord.SetEasting(Row.Cells(EastingCol).Value, Coordinate.UpdateMode.None) 'Update the new Input Easting value
                Conversion.InputCrs.Coord.SetNorthing(Row.Cells(NorthingCol).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Northing value
                UpdateDatumTransTable(Row.Index)
            Next

        ElseIf rbEnterInputLongLat.Checked Then
            Dim LongitudeCol As Integer
            Dim LatitudeCol As Integer
            Dim HeightCol As Integer
            Dim RowNo As Integer
            For Each Col As DataGridViewColumn In dgvConversion.Columns
                If Col.HeaderText = "Input Longitude" Then LongitudeCol = Col.Index
                If Col.HeaderText = "Input Latitude" Then LatitudeCol = Col.Index
                If Col.HeaderText = "Input Ellipsoidal Height" Then HeightCol = Col.Index
            Next
            For Each Row As DataGridViewRow In dgvConversion.Rows
                Conversion.InputCrs.Coord.SetLongitude(Row.Cells(LongitudeCol).Value, Coordinate.UpdateMode.None) 'Update the new Input Longitude value
                Conversion.InputCrs.Coord.SetLatitude(Row.Cells(LatitudeCol).Value, Coordinate.UpdateMode.None) 'Update the new Input Latitude value
                Conversion.InputCrs.Coord.SetEllipsoidalHeight(Row.Cells(HeightCol).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Ellipsoidal Height value
                UpdateDatumTransTable(Row.Index)
            Next


        ElseIf rbEnterInputXYZ.Checked Then
            Dim XCol As Integer
            Dim YCol As Integer
            Dim ZCol As Integer
            Dim RowNo As Integer
            For Each Col As DataGridViewColumn In dgvConversion.Columns
                If Col.HeaderText = "Input X" Then XCol = Col.Index
                If Col.HeaderText = "Input Y" Then YCol = Col.Index
                If Col.HeaderText = "Input Z" Then ZCol = Col.Index
            Next
            For Each Row As DataGridViewRow In dgvConversion.Rows
                Conversion.InputCrs.Coord.SetLongitude(Row.Cells(XCol).Value, Coordinate.UpdateMode.None) 'Update the new Input X value
                Conversion.InputCrs.Coord.SetLatitude(Row.Cells(YCol).Value, Coordinate.UpdateMode.None) 'Update the new Input Y value
                Conversion.InputCrs.Coord.SetEllipsoidalHeight(Row.Cells(ZCol).Value, Coordinate.UpdateMode.InputOutputAll) 'Update the new Input Z value
                UpdateDatumTransTable(Row.Index)
            Next

        End If

        dgvConversion.AllowUserToAddRows = True

    End Sub


    Private Sub UpdateOutputGeographicCoords()
        'Update all of the Output Grographic Coordinates.
        'This is done if the Geographic display format has changed.

        If rbDecDegrees.Checked Then
            Try
                txtDegreeDecPlaces.Text = Int(txtDegreeDecPlaces.Text.Trim)
                Dim Format As String = "F" & txtDegreeDecPlaces.Text.Trim
                dgvConversion.Columns(5).DefaultCellStyle.Format = Format
                dgvConversion.Columns(6).DefaultCellStyle.Format = Format
                dgvConversion.Columns(17).DefaultCellStyle.Format = Format
                dgvConversion.Columns(18).DefaultCellStyle.Format = Format
            Catch ex As Exception

            End Try
        End If


        dgvConversion.AllowUserToAddRows = False

        If Conversion.InputCrs.Kind = CoordRefSystem.CrsKind.projected Then
            rbEnterInputEastNorth.Enabled = True
        Else
            rbEnterInputEastNorth.Enabled = False
            If rbEnterInputEastNorth.Checked Then rbEnterInputLongLat.Checked = True
        End If

        If rbEnterInputEastNorth.Checked Then
            For Each Row As DataGridViewRow In dgvConversion.Rows
                Conversion.InputCrs.Coord.SetEasting(Row.Cells(3).Value, Coordinate.UpdateMode.None)
                Conversion.InputCrs.Coord.SetNorthing(Row.Cells(4).Value, Coordinate.UpdateMode.TransLongLat)
                If rbDecDegrees.Checked Then
                    Row.Cells(17).Value = Conversion.OutputCrs.Coord.Longitude
                    Row.Cells(18).Value = Conversion.OutputCrs.Coord.Latitude
                Else
                    Row.Cells(17).Value = Conversion.OutputCrs.Coord.LongitudeDMS
                    Row.Cells(18).Value = Conversion.OutputCrs.Coord.LatitudeDMS
                End If
            Next

        ElseIf rbEnterInputLongLat.Checked Then
            For Each Row As DataGridViewRow In dgvConversion.Rows
                Conversion.InputCrs.Coord.SetLongitude(Row.Cells(5).Value, Coordinate.UpdateMode.None)
                Conversion.InputCrs.Coord.SetLatitude(Row.Cells(6).Value, Coordinate.UpdateMode.None)
                Conversion.InputCrs.Coord.SetEllipsoidalHeight(Row.Cells(7).Value, Coordinate.UpdateMode.TransLongLat)
                If rbDecDegrees.Checked Then
                    Row.Cells(17).Value = Conversion.OutputCrs.Coord.Longitude
                    Row.Cells(18).Value = Conversion.OutputCrs.Coord.Latitude
                Else
                    Row.Cells(17).Value = Conversion.OutputCrs.Coord.LongitudeDMS
                    Row.Cells(18).Value = Conversion.OutputCrs.Coord.LatitudeDMS
                End If
            Next

        ElseIf rbEnterInputXYZ.Checked Then
            For Each Row As DataGridViewRow In dgvConversion.Rows
                Conversion.InputCrs.Coord.SetX(Row.Cells(8).Value, Coordinate.UpdateMode.None)
                Conversion.InputCrs.Coord.SetY(Row.Cells(9).Value, Coordinate.UpdateMode.None)
                Conversion.InputCrs.Coord.SetZ(Row.Cells(10).Value, Coordinate.UpdateMode.TransLongLat)
                If rbDecDegrees.Checked Then
                    Row.Cells(17).Value = Conversion.OutputCrs.Coord.Longitude
                    Row.Cells(18).Value = Conversion.OutputCrs.Coord.Latitude
                Else
                    Row.Cells(17).Value = Conversion.OutputCrs.Coord.LongitudeDMS
                    Row.Cells(18).Value = Conversion.OutputCrs.Coord.LatitudeDMS
                End If
            Next

        Else

        End If

        dgvConversion.AllowUserToAddRows = True
        dgvConversion.AutoResizeColumns()

    End Sub





#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Events - Events that can be triggered by this form." '==========================================================================================================================

#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class