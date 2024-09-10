<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblRows = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblProcessingRow = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pbrProgress = New System.Windows.Forms.ToolStripProgressBar()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InfoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tspMenu = New System.Windows.Forms.ToolStrip()
        Me.butProcess = New System.Windows.Forms.ToolStripButton()
        Me.butExport = New System.Windows.Forms.ToolStripButton()
        Me.butRefresh = New System.Windows.Forms.ToolStripButton()
        Me.butSettings = New System.Windows.Forms.ToolStripButton()
        Me.butExit = New System.Windows.Forms.ToolStripButton()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboOrderPoint = New System.Windows.Forms.ComboBox()
        Me.cboReportType = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboSortBy = New System.Windows.Forms.ComboBox()
        Me.cboStore = New System.Windows.Forms.ComboBox()
        Me.chkProcessExport = New System.Windows.Forms.CheckBox()
        Me.lsvFiles = New System.Windows.Forms.ListView()
        Me.colName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colPath = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.dgvData = New System.Windows.Forms.DataGridView()
        Me.bgWorker = New System.ComponentModel.BackgroundWorker()
        Me.StatusStrip1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.tspMenu.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.lblRows, Me.lblProcessingRow, Me.pbrProgress})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 493)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(824, 22)
        Me.StatusStrip1.TabIndex = 0
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(110, 17)
        Me.ToolStripStatusLabel1.Text = "Updated: 9/10/2024"
        '
        'lblRows
        '
        Me.lblRows.Name = "lblRows"
        Me.lblRows.Size = New System.Drawing.Size(66, 17)
        Me.lblRows.Text = "Total Rows:"
        Me.lblRows.Visible = False
        '
        'lblProcessingRow
        '
        Me.lblProcessingRow.Name = "lblProcessingRow"
        Me.lblProcessingRow.Size = New System.Drawing.Size(0, 17)
        '
        'pbrProgress
        '
        Me.pbrProgress.Name = "pbrProgress"
        Me.pbrProgress.Size = New System.Drawing.Size(100, 16)
        Me.pbrProgress.Visible = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(824, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(93, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.InfoToolStripMenuItem})
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'InfoToolStripMenuItem
        '
        Me.InfoToolStripMenuItem.Name = "InfoToolStripMenuItem"
        Me.InfoToolStripMenuItem.Size = New System.Drawing.Size(95, 22)
        Me.InfoToolStripMenuItem.Text = "Info"
        '
        'tspMenu
        '
        Me.tspMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.butProcess, Me.butExport, Me.butRefresh, Me.butSettings, Me.butExit})
        Me.tspMenu.Location = New System.Drawing.Point(0, 24)
        Me.tspMenu.Name = "tspMenu"
        Me.tspMenu.Size = New System.Drawing.Size(824, 54)
        Me.tspMenu.TabIndex = 2
        Me.tspMenu.Text = "ToolStrip1"
        '
        'butProcess
        '
        Me.butProcess.Enabled = False
        Me.butProcess.Image = CType(resources.GetObject("butProcess.Image"), System.Drawing.Image)
        Me.butProcess.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butProcess.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butProcess.Name = "butProcess"
        Me.butProcess.Size = New System.Drawing.Size(51, 51)
        Me.butProcess.Text = "Process"
        Me.butProcess.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'butExport
        '
        Me.butExport.Enabled = False
        Me.butExport.Image = CType(resources.GetObject("butExport.Image"), System.Drawing.Image)
        Me.butExport.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butExport.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butExport.Name = "butExport"
        Me.butExport.Size = New System.Drawing.Size(45, 51)
        Me.butExport.Text = "Export"
        Me.butExport.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'butRefresh
        '
        Me.butRefresh.Image = CType(resources.GetObject("butRefresh.Image"), System.Drawing.Image)
        Me.butRefresh.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butRefresh.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butRefresh.Name = "butRefresh"
        Me.butRefresh.Size = New System.Drawing.Size(50, 51)
        Me.butRefresh.Text = "Refresh"
        Me.butRefresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.butRefresh.ToolTipText = "Refresh file list"
        '
        'butSettings
        '
        Me.butSettings.Image = CType(resources.GetObject("butSettings.Image"), System.Drawing.Image)
        Me.butSettings.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butSettings.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butSettings.Name = "butSettings"
        Me.butSettings.Size = New System.Drawing.Size(53, 51)
        Me.butSettings.Text = "Settings"
        Me.butSettings.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.butSettings.Visible = False
        '
        'butExit
        '
        Me.butExit.Image = CType(resources.GetObject("butExit.Image"), System.Drawing.Image)
        Me.butExit.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butExit.Name = "butExit"
        Me.butExit.Size = New System.Drawing.Size(36, 51)
        Me.butExit.Text = "Exit"
        Me.butExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 78)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.lsvFiles)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvData)
        Me.SplitContainer1.Size = New System.Drawing.Size(800, 415)
        Me.SplitContainer1.SplitterDistance = 225
        Me.SplitContainer1.TabIndex = 3
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cboOrderPoint)
        Me.GroupBox1.Controls.Add(Me.cboReportType)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.cboSortBy)
        Me.GroupBox1.Controls.Add(Me.cboStore)
        Me.GroupBox1.Controls.Add(Me.chkProcessExport)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(222, 166)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Report Options"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Report Type"
        '
        'cboOrderPoint
        '
        Me.cboOrderPoint.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboOrderPoint.Enabled = False
        Me.cboOrderPoint.FormattingEnabled = True
        Me.cboOrderPoint.Items.AddRange(New Object() {"Average Sales", "Average Sales by 6 mo.", "Suggested Order", "Suggested Order by 6 mo.", "Suggested Order by 11 mo.", "None", "Zero"})
        Me.cboOrderPoint.Location = New System.Drawing.Point(81, 82)
        Me.cboOrderPoint.Name = "cboOrderPoint"
        Me.cboOrderPoint.Size = New System.Drawing.Size(133, 21)
        Me.cboOrderPoint.TabIndex = 12
        '
        'cboReportType
        '
        Me.cboReportType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReportType.FormattingEnabled = True
        Me.cboReportType.Items.AddRange(New Object() {"Overstock/UndertStock", "Z-Report - Dept", "Z-Report - Vendor"})
        Me.cboReportType.Location = New System.Drawing.Point(81, 27)
        Me.cboReportType.Name = "cboReportType"
        Me.cboReportType.Size = New System.Drawing.Size(133, 21)
        Me.cboReportType.Sorted = True
        Me.cboReportType.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 86)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Order Point"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Store No."
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 138)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(83, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Process/Export:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(44, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Sort By:"
        '
        'cboSortBy
        '
        Me.cboSortBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSortBy.FormattingEnabled = True
        Me.cboSortBy.Items.AddRange(New Object() {"Item Number", "Dept and Class"})
        Me.cboSortBy.Location = New System.Drawing.Point(81, 109)
        Me.cboSortBy.Name = "cboSortBy"
        Me.cboSortBy.Size = New System.Drawing.Size(133, 21)
        Me.cboSortBy.TabIndex = 8
        '
        'cboStore
        '
        Me.cboStore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStore.Enabled = False
        Me.cboStore.FormattingEnabled = True
        Me.cboStore.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9"})
        Me.cboStore.Location = New System.Drawing.Point(81, 55)
        Me.cboStore.Name = "cboStore"
        Me.cboStore.Size = New System.Drawing.Size(133, 21)
        Me.cboStore.TabIndex = 5
        '
        'chkProcessExport
        '
        Me.chkProcessExport.AutoSize = True
        Me.chkProcessExport.Location = New System.Drawing.Point(95, 137)
        Me.chkProcessExport.Name = "chkProcessExport"
        Me.chkProcessExport.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkProcessExport.Size = New System.Drawing.Size(15, 14)
        Me.chkProcessExport.TabIndex = 7
        Me.chkProcessExport.UseVisualStyleBackColor = True
        '
        'lsvFiles
        '
        Me.lsvFiles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lsvFiles.BackColor = System.Drawing.SystemColors.Window
        Me.lsvFiles.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colName, Me.colPath})
        Me.lsvFiles.Enabled = False
        Me.lsvFiles.FullRowSelect = True
        Me.lsvFiles.GridLines = True
        Me.lsvFiles.HideSelection = False
        Me.lsvFiles.Location = New System.Drawing.Point(0, 175)
        Me.lsvFiles.Name = "lsvFiles"
        Me.lsvFiles.Size = New System.Drawing.Size(222, 237)
        Me.lsvFiles.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lsvFiles.TabIndex = 0
        Me.lsvFiles.UseCompatibleStateImageBehavior = False
        Me.lsvFiles.View = System.Windows.Forms.View.Details
        '
        'colName
        '
        Me.colName.Text = "Filename"
        Me.colName.Width = 220
        '
        'colPath
        '
        Me.colPath.Width = 120
        '
        'dgvData
        '
        Me.dgvData.AllowUserToAddRows = False
        Me.dgvData.AllowUserToDeleteRows = False
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.GradientActiveCaption
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvData.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvData.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvData.Location = New System.Drawing.Point(0, 0)
        Me.dgvData.MultiSelect = False
        Me.dgvData.Name = "dgvData"
        Me.dgvData.ReadOnly = True
        Me.dgvData.Size = New System.Drawing.Size(571, 412)
        Me.dgvData.TabIndex = 1
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(824, 515)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.tspMenu)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmMain"
        Me.Text = "Eagle Report Tool (Rapid) - Surangel And Sons"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.tspMenu.ResumeLayout(False)
        Me.tspMenu.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents tspMenu As ToolStrip
    Friend WithEvents butProcess As ToolStripButton
    Friend WithEvents butExport As ToolStripButton
    Friend WithEvents butSettings As ToolStripButton
    Friend WithEvents butExit As ToolStripButton
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents lsvFiles As ListView
    Friend WithEvents colName As ColumnHeader
    Friend WithEvents colPath As ColumnHeader
    Friend WithEvents dgvData As DataGridView
    Friend WithEvents lblRows As ToolStripStatusLabel
    Friend WithEvents lblProcessingRow As ToolStripStatusLabel
    Friend WithEvents pbrProgress As ToolStripProgressBar
    Friend WithEvents InfoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents cboReportType As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents cboStore As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents butRefresh As ToolStripButton
    Friend WithEvents Label3 As Label
    Friend WithEvents chkProcessExport As CheckBox
    Friend WithEvents cboSortBy As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents cboOrderPoint As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents bgWorker As System.ComponentModel.BackgroundWorker
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
End Class
