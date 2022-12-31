<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCoordOpMethods
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.txtFind = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.txtQuery = New System.Windows.Forms.TextBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.dgvMethodList = New System.Windows.Forms.DataGridView()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.udRowNo = New System.Windows.Forms.NumericUpDown()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.txtMethodFormula = New System.Windows.Forms.TextBox()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.txtMethodExample = New System.Windows.Forms.TextBox()
        Me.txtMethodRemarks = New System.Windows.Forms.TextBox()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.txtMethodCode = New System.Windows.Forms.TextBox()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.udFont = New System.Windows.Forms.NumericUpDown()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.txtMethodReversable = New System.Windows.Forms.TextBox()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.txtMethodName = New System.Windows.Forms.TextBox()
        Me.txtNInputRecords = New System.Windows.Forms.TextBox()
        Me.Label143 = New System.Windows.Forms.Label()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvMethodList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.udRowNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.udFont, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(1509, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 40)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1545, 826)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.txtNInputRecords)
        Me.TabPage1.Controls.Add(Me.Label143)
        Me.TabPage1.Controls.Add(Me.btnSelect)
        Me.TabPage1.Controls.Add(Me.btnFind)
        Me.TabPage1.Controls.Add(Me.txtFind)
        Me.TabPage1.Controls.Add(Me.Label26)
        Me.TabPage1.Controls.Add(Me.btnApply)
        Me.TabPage1.Controls.Add(Me.txtQuery)
        Me.TabPage1.Controls.Add(Me.Label25)
        Me.TabPage1.Controls.Add(Me.dgvMethodList)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1537, 800)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Search"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(6, 34)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(48, 22)
        Me.btnSelect.TabIndex = 22
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(317, 32)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(48, 22)
        Me.btnFind.TabIndex = 21
        Me.btnFind.Text = "Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'txtFind
        '
        Me.txtFind.Location = New System.Drawing.Point(167, 32)
        Me.txtFind.Name = "txtFind"
        Me.txtFind.Size = New System.Drawing.Size(144, 20)
        Me.txtFind.TabIndex = 20
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(80, 35)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(81, 13)
        Me.Label26.TabIndex = 19
        Me.Label26.Text = "Name contains:"
        '
        'btnApply
        '
        Me.btnApply.Location = New System.Drawing.Point(6, 4)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(48, 22)
        Me.btnApply.TabIndex = 18
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'txtQuery
        '
        Me.txtQuery.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQuery.Location = New System.Drawing.Point(104, 6)
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(1427, 20)
        Me.txtQuery.TabIndex = 17
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(60, 9)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(38, 13)
        Me.Label25.TabIndex = 16
        Me.Label25.Text = "Query:"
        '
        'dgvMethodList
        '
        Me.dgvMethodList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMethodList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMethodList.Location = New System.Drawing.Point(6, 60)
        Me.dgvMethodList.Name = "dgvMethodList"
        Me.dgvMethodList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvMethodList.Size = New System.Drawing.Size(1525, 734)
        Me.dgvMethodList.TabIndex = 15
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.udRowNo)
        Me.TabPage2.Controls.Add(Me.SplitContainer1)
        Me.TabPage2.Controls.Add(Me.txtMethodRemarks)
        Me.TabPage2.Controls.Add(Me.Label46)
        Me.TabPage2.Controls.Add(Me.txtMethodCode)
        Me.TabPage2.Controls.Add(Me.Label45)
        Me.TabPage2.Controls.Add(Me.udFont)
        Me.TabPage2.Controls.Add(Me.Label44)
        Me.TabPage2.Controls.Add(Me.txtMethodReversable)
        Me.TabPage2.Controls.Add(Me.Label43)
        Me.TabPage2.Controls.Add(Me.Label40)
        Me.TabPage2.Controls.Add(Me.txtMethodName)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1537, 800)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Coordinate Operation Method"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'udRowNo
        '
        Me.udRowNo.Location = New System.Drawing.Point(160, 32)
        Me.udRowNo.Name = "udRowNo"
        Me.udRowNo.Size = New System.Drawing.Size(65, 20)
        Me.udRowNo.TabIndex = 60
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 58)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label41)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtMethodFormula)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label42)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtMethodExample)
        Me.SplitContainer1.Size = New System.Drawing.Size(1531, 736)
        Me.SplitContainer1.SplitterDistance = 597
        Me.SplitContainer1.TabIndex = 59
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(3, 0)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(47, 13)
        Me.Label41.TabIndex = 33
        Me.Label41.Text = "Formula:"
        '
        'txtMethodFormula
        '
        Me.txtMethodFormula.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMethodFormula.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMethodFormula.Location = New System.Drawing.Point(3, 16)
        Me.txtMethodFormula.Multiline = True
        Me.txtMethodFormula.Name = "txtMethodFormula"
        Me.txtMethodFormula.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMethodFormula.Size = New System.Drawing.Size(591, 717)
        Me.txtMethodFormula.TabIndex = 32
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(3, 0)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(50, 13)
        Me.Label42.TabIndex = 34
        Me.Label42.Text = "Example:"
        '
        'txtMethodExample
        '
        Me.txtMethodExample.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMethodExample.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMethodExample.Location = New System.Drawing.Point(3, 16)
        Me.txtMethodExample.Multiline = True
        Me.txtMethodExample.Name = "txtMethodExample"
        Me.txtMethodExample.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMethodExample.Size = New System.Drawing.Size(924, 717)
        Me.txtMethodExample.TabIndex = 35
        '
        'txtMethodRemarks
        '
        Me.txtMethodRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMethodRemarks.Location = New System.Drawing.Point(561, 6)
        Me.txtMethodRemarks.Multiline = True
        Me.txtMethodRemarks.Name = "txtMethodRemarks"
        Me.txtMethodRemarks.Size = New System.Drawing.Size(970, 46)
        Me.txtMethodRemarks.TabIndex = 57
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Location = New System.Drawing.Point(499, 9)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(52, 13)
        Me.Label46.TabIndex = 56
        Me.Label46.Text = "Remarks:"
        '
        'txtMethodCode
        '
        Me.txtMethodCode.Location = New System.Drawing.Point(50, 32)
        Me.txtMethodCode.Name = "txtMethodCode"
        Me.txtMethodCode.Size = New System.Drawing.Size(104, 20)
        Me.txtMethodCode.TabIndex = 55
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Location = New System.Drawing.Point(9, 34)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(35, 13)
        Me.Label45.TabIndex = 54
        Me.Label45.Text = "Code:"
        '
        'udFont
        '
        Me.udFont.Location = New System.Drawing.Point(432, 32)
        Me.udFont.Name = "udFont"
        Me.udFont.Size = New System.Drawing.Size(72, 20)
        Me.udFont.TabIndex = 53
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(374, 36)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(52, 13)
        Me.Label44.TabIndex = 52
        Me.Label44.Text = "Font size:"
        '
        'txtMethodReversable
        '
        Me.txtMethodReversable.Location = New System.Drawing.Point(301, 32)
        Me.txtMethodReversable.Name = "txtMethodReversable"
        Me.txtMethodReversable.Size = New System.Drawing.Size(67, 20)
        Me.txtMethodReversable.TabIndex = 51
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(231, 35)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(64, 13)
        Me.Label43.TabIndex = 50
        Me.Label43.Text = "Reversable:"
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(6, 9)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(75, 13)
        Me.Label40.TabIndex = 48
        Me.Label40.Text = "Method name:"
        '
        'txtMethodName
        '
        Me.txtMethodName.Location = New System.Drawing.Point(87, 6)
        Me.txtMethodName.Name = "txtMethodName"
        Me.txtMethodName.Size = New System.Drawing.Size(406, 20)
        Me.txtMethodName.TabIndex = 49
        '
        'txtNInputRecords
        '
        Me.txtNInputRecords.Location = New System.Drawing.Point(442, 32)
        Me.txtNInputRecords.Name = "txtNInputRecords"
        Me.txtNInputRecords.ReadOnly = True
        Me.txtNInputRecords.Size = New System.Drawing.Size(127, 20)
        Me.txtNInputRecords.TabIndex = 24
        '
        'Label143
        '
        Me.Label143.AutoSize = True
        Me.Label143.Location = New System.Drawing.Point(371, 35)
        Me.Label143.Name = "Label143"
        Me.Label143.Size = New System.Drawing.Size(65, 13)
        Me.Label143.TabIndex = 23
        Me.Label143.Text = "No. records:"
        '
        'frmCoordOpMethods
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1569, 878)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmCoordOpMethods"
        Me.Text = "Coordinate Operation Methods"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.dgvMethodList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.udRowNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.udFont, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents btnSelect As Button
    Friend WithEvents btnFind As Button
    Friend WithEvents txtFind As TextBox
    Friend WithEvents Label26 As Label
    Friend WithEvents btnApply As Button
    Friend WithEvents txtQuery As TextBox
    Friend WithEvents Label25 As Label
    Friend WithEvents dgvMethodList As DataGridView
    Friend WithEvents txtMethodRemarks As TextBox
    Friend WithEvents Label46 As Label
    Friend WithEvents txtMethodCode As TextBox
    Friend WithEvents Label45 As Label
    Friend WithEvents udFont As NumericUpDown
    Friend WithEvents Label44 As Label
    Friend WithEvents txtMethodReversable As TextBox
    Friend WithEvents Label43 As Label
    Friend WithEvents Label40 As Label
    Friend WithEvents txtMethodName As TextBox
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Label41 As Label
    Friend WithEvents txtMethodFormula As TextBox
    Friend WithEvents Label42 As Label
    Friend WithEvents txtMethodExample As TextBox
    Friend WithEvents udRowNo As NumericUpDown
    Friend WithEvents txtNInputRecords As TextBox
    Friend WithEvents Label143 As Label
End Class
