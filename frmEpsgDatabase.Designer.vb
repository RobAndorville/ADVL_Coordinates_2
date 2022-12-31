<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEpsgDatabase
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
        Me.btnExit = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btnApplyQuery = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtQuery = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbSelectTable = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDatabase = New System.Windows.Forms.TextBox()
        Me.btnFindDatabase = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtNRecords = New System.Windows.Forms.TextBox()
        Me.rbViewAll = New System.Windows.Forms.RadioButton()
        Me.rbViewFirst = New System.Windows.Forms.RadioButton()
        Me.txtViewNRecords = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(1070, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 19
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 177)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1122, 692)
        Me.DataGridView1.TabIndex = 39
        '
        'btnApplyQuery
        '
        Me.btnApplyQuery.Location = New System.Drawing.Point(12, 119)
        Me.btnApplyQuery.Name = "btnApplyQuery"
        Me.btnApplyQuery.Size = New System.Drawing.Size(47, 22)
        Me.btnApplyQuery.TabIndex = 38
        Me.btnApplyQuery.Text = "Apply"
        Me.btnApplyQuery.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 98)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 37
        Me.Label2.Text = "Query:"
        '
        'txtQuery
        '
        Me.txtQuery.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQuery.Location = New System.Drawing.Point(84, 95)
        Me.txtQuery.Multiline = True
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(1050, 46)
        Me.txtQuery.TabIndex = 36
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 71)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "Select table:"
        '
        'cmbSelectTable
        '
        Me.cmbSelectTable.FormattingEnabled = True
        Me.cmbSelectTable.Location = New System.Drawing.Point(84, 68)
        Me.cmbSelectTable.Name = "cmbSelectTable"
        Me.cmbSelectTable.Size = New System.Drawing.Size(269, 21)
        Me.cmbSelectTable.TabIndex = 34
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 44)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 13)
        Me.Label3.TabIndex = 33
        Me.Label3.Text = "Database:"
        '
        'txtDatabase
        '
        Me.txtDatabase.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDatabase.Location = New System.Drawing.Point(84, 41)
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(995, 20)
        Me.txtDatabase.TabIndex = 32
        '
        'btnFindDatabase
        '
        Me.btnFindDatabase.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindDatabase.Location = New System.Drawing.Point(1085, 40)
        Me.btnFindDatabase.Name = "btnFindDatabase"
        Me.btnFindDatabase.Size = New System.Drawing.Size(49, 22)
        Me.btnFindDatabase.TabIndex = 31
        Me.btnFindDatabase.Text = "Find"
        Me.btnFindDatabase.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 150)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(61, 13)
        Me.Label4.TabIndex = 40
        Me.Label4.Text = "N Records:"
        '
        'txtNRecords
        '
        Me.txtNRecords.Location = New System.Drawing.Point(84, 147)
        Me.txtNRecords.Name = "txtNRecords"
        Me.txtNRecords.Size = New System.Drawing.Size(269, 20)
        Me.txtNRecords.TabIndex = 41
        '
        'rbViewAll
        '
        Me.rbViewAll.AutoSize = True
        Me.rbViewAll.Location = New System.Drawing.Point(359, 69)
        Me.rbViewAll.Name = "rbViewAll"
        Me.rbViewAll.Size = New System.Drawing.Size(61, 17)
        Me.rbViewAll.TabIndex = 42
        Me.rbViewAll.TabStop = True
        Me.rbViewAll.Text = "View all"
        Me.rbViewAll.UseVisualStyleBackColor = True
        '
        'rbViewFirst
        '
        Me.rbViewFirst.AutoSize = True
        Me.rbViewFirst.Location = New System.Drawing.Point(427, 69)
        Me.rbViewFirst.Name = "rbViewFirst"
        Me.rbViewFirst.Size = New System.Drawing.Size(67, 17)
        Me.rbViewFirst.TabIndex = 43
        Me.rbViewFirst.TabStop = True
        Me.rbViewFirst.Text = "View first"
        Me.rbViewFirst.UseVisualStyleBackColor = True
        '
        'txtViewNRecords
        '
        Me.txtViewNRecords.Location = New System.Drawing.Point(503, 69)
        Me.txtViewNRecords.Name = "txtViewNRecords"
        Me.txtViewNRecords.Size = New System.Drawing.Size(121, 20)
        Me.txtViewNRecords.TabIndex = 44
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(630, 71)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(42, 13)
        Me.Label5.TabIndex = 45
        Me.Label5.Text = "records"
        '
        'frmEpsgDatabase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1146, 882)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtViewNRecords)
        Me.Controls.Add(Me.rbViewFirst)
        Me.Controls.Add(Me.rbViewAll)
        Me.Controls.Add(Me.txtNRecords)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btnApplyQuery)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtQuery)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbSelectTable)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDatabase)
        Me.Controls.Add(Me.btnFindDatabase)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmEpsgDatabase"
        Me.Text = "EPSG Database"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents btnApplyQuery As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents txtQuery As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbSelectTable As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtDatabase As TextBox
    Friend WithEvents btnFindDatabase As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Label4 As Label
    Friend WithEvents txtNRecords As TextBox
    Friend WithEvents rbViewAll As RadioButton
    Friend WithEvents rbViewFirst As RadioButton
    Friend WithEvents txtViewNRecords As TextBox
    Friend WithEvents Label5 As Label
End Class
