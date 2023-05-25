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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.lblRows = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblProcessingRow = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pbrProgress = New System.Windows.Forms.ToolStripProgressBar()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InfoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.butProcess = New System.Windows.Forms.ToolStripButton()
        Me.butExport = New System.Windows.Forms.ToolStripButton()
        Me.butSettings = New System.Windows.Forms.ToolStripButton()
        Me.butExit = New System.Windows.Forms.ToolStripButton()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lsvFiles = New System.Windows.Forms.ListView()
        Me.colName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colPath = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.dgvData = New System.Windows.Forms.DataGridView()
        Me.StatusStrip1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblRows, Me.lblProcessingRow, Me.pbrProgress})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 428)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(824, 22)
        Me.StatusStrip1.TabIndex = 0
        Me.StatusStrip1.Text = "StatusStrip1"
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
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.butProcess, Me.butExport, Me.butSettings, Me.butExit})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 24)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(824, 54)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
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
        'butSettings
        '
        Me.butSettings.Image = CType(resources.GetObject("butSettings.Image"), System.Drawing.Image)
        Me.butSettings.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.butSettings.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butSettings.Name = "butSettings"
        Me.butSettings.Size = New System.Drawing.Size(53, 51)
        Me.butSettings.Text = "Settings"
        Me.butSettings.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.lsvFiles)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvData)
        Me.SplitContainer1.Size = New System.Drawing.Size(800, 350)
        Me.SplitContainer1.SplitterDistance = 225
        Me.SplitContainer1.TabIndex = 3
        '
        'lsvFiles
        '
        Me.lsvFiles.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colName, Me.colPath})
        Me.lsvFiles.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lsvFiles.Enabled = False
        Me.lsvFiles.FullRowSelect = True
        Me.lsvFiles.GridLines = True
        Me.lsvFiles.HideSelection = False
        Me.lsvFiles.Location = New System.Drawing.Point(0, 0)
        Me.lsvFiles.MultiSelect = False
        Me.lsvFiles.Name = "lsvFiles"
        Me.lsvFiles.Size = New System.Drawing.Size(225, 350)
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
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvData.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvData.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvData.Location = New System.Drawing.Point(0, 0)
        Me.dgvData.MultiSelect = False
        Me.dgvData.Name = "dgvData"
        Me.dgvData.ReadOnly = True
        Me.dgvData.Size = New System.Drawing.Size(571, 350)
        Me.dgvData.TabIndex = 1
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(824, 450)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
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
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStrip1 As ToolStrip
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
End Class
