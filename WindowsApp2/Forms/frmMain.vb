Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices

Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Common
Imports System.Windows.Input
Imports System.Data.SqlTypes
Imports System.Globalization
Imports System.Runtime.Serialization.Formatters
Imports System.Security.Cryptography
Imports System.Windows.Forms.AxHost
'Imports System.Runtime.InteropServices

Public Class frmMain

    Dim Thread As System.Threading.Thread

    'Database Connection
    Private xlConn As New OleDb.OleDbConnection
    Private xlDataset As New System.Data.DataSet
    Private xlDataAdapter As New System.Data.OleDb.OleDbDataAdapter
    Private xlCmd As New System.Data.OleDb.OleDbCommand

    Private DirPathEpicor As String = "C:\Users\" & Environment.UserName & "\AppData\Local\Temp\Epicor"
    Private DirPath3Apps As String = "C:\3apps\temp"
    Private strExcel_Path As String
    Private strExcel_FileName As String
    Private strVendorName As String
    Private intReportId As Integer
    Private intStoreId As Integer

    Private isFormLoaded As Boolean
    Private isExcelFileLoaded As Boolean
    Private isPopMessage As Boolean = False
    Private isMultiSelect As Boolean
    Private isSuggestedOrderOn As Boolean
    'Private intSuggestedOrder As Integer

    Private blnCancel As Boolean = False

    Private Sub DataGridView_InsertColumns(intReportId As Integer)
        Dim strStore As New DataGridViewTextBoxColumn()
        Dim colQOH As New DataGridViewTextBoxColumn()
        Dim colQOO As New DataGridViewTextBoxColumn()
        Dim colQTY_AVAIL As New DataGridViewTextBoxColumn()
        Dim colORDER_POINT As New DataGridViewTextBoxColumn()
        Dim colTOTAL As New DataGridViewTextBoxColumn()
        Dim colALT_VENDOR As New DataGridViewTextBoxColumn()

        Dim colPERIOD1 As New DataGridViewTextBoxColumn()
        Dim colPERIOD2 As New DataGridViewTextBoxColumn()
        Dim colPERIOD3 As New DataGridViewTextBoxColumn()
        Dim colPERIOD4 As New DataGridViewTextBoxColumn()
        Dim colPERIOD5 As New DataGridViewTextBoxColumn()
        Dim colPERIOD6 As New DataGridViewTextBoxColumn()
        Dim colPERIOD7 As New DataGridViewTextBoxColumn()
        Dim colPERIOD8 As New DataGridViewTextBoxColumn()
        Dim colPERIOD9 As New DataGridViewTextBoxColumn()
        Dim colPERIOD10 As New DataGridViewTextBoxColumn()
        Dim colPERIOD11 As New DataGridViewTextBoxColumn()
        Dim colPERIOD12 As New DataGridViewTextBoxColumn()

        Dim colItemNumber1 As New DataGridViewTextBoxColumn()
        Dim colItemNumber2 As New DataGridViewTextBoxColumn()
        Dim colItemNumberLength As New DataGridViewTextBoxColumn()
        Dim colItemNumberFullLength As New DataGridViewTextBoxColumn()

        'Added 9/9/2024
        'Dim colQtyByStore As New DataGridViewTextBoxColumn()

        'Added 9/10/2024
        Dim st1Qty As New DataGridViewTextBoxColumn()
        Dim st2Qty As New DataGridViewTextBoxColumn()
        Dim st3Qty As New DataGridViewTextBoxColumn()
        Dim st4Qty As New DataGridViewTextBoxColumn()
        Dim st6Qty As New DataGridViewTextBoxColumn()
        Dim st7Qty As New DataGridViewTextBoxColumn()
        Dim st9Qty As New DataGridViewTextBoxColumn()


        'UO REPORT
        Dim colSuggestedOrder As New DataGridViewTextBoxColumn()
        Dim colYTDUnits As New DataGridViewTextBoxColumn()
        Dim colLYRUnits As New DataGridViewTextBoxColumn()
        Dim colAvgCost As New DataGridViewTextBoxColumn()
        Dim colPopCode As New DataGridViewTextBoxColumn()
        Dim colDateLastSale As New DataGridViewTextBoxColumn()
        Dim colDateLastReceipt As New DataGridViewTextBoxColumn()
        Dim colTurnsPlus As New DataGridViewTextBoxColumn()
        Dim colStoreCloseout As New DataGridViewTextBoxColumn()
        Dim colDiscontinued As New DataGridViewTextBoxColumn()
        Dim colCurrSalesUnits As New DataGridViewTextBoxColumn()
        Dim colRetail As New DataGridViewTextBoxColumn()
        Dim colUnusedS217 As New DataGridViewTextBoxColumn()


        With strStore
            .Name = "STORE"
            .HeaderText = "STORE"
        End With

        With colQOH
            .Name = "QOH"
            .HeaderText = "QOH"
        End With

        With colQOO
            .Name = "QOO"
            .HeaderText = "QOO"
        End With

        With colQTY_AVAIL
            .Name = "QTY_AVAIL"
            .HeaderText = "QTY AVAILABLE"
        End With

        With colORDER_POINT
            .Name = "ORDER_POINT"
            .HeaderText = "ORDER POINT"
        End With

        With colTOTAL
            .Name = "TOTAL"
            .HeaderText = "TOTAL"
        End With

        With colPERIOD1
            .Name = "PERIOD1"
            .HeaderText = "PERIOD 1"
            .Width = 50
        End With

        With colPERIOD2
            .Name = "PERIOD2"
            .HeaderText = "PERIOD 2"
            .Width = 50
        End With

        With colPERIOD3
            .Name = "PERIOD3"
            .HeaderText = "PERIOD 3"
            .Width = 50
        End With

        With colPERIOD4
            .Name = "PERIOD4"
            .HeaderText = "PERIOD 4"
            .Width = 50
        End With

        With colPERIOD5
            .Name = "PERIOD5"
            .HeaderText = "PERIOD 5"
            .Width = 50
        End With

        With colPERIOD6
            .Name = "PERIOD6"
            .HeaderText = "PERIOD 6"
            .Width = 50
        End With

        With colPERIOD7
            .Name = "PERIOD7"
            .HeaderText = "PERIOD 7"
            .Width = 50
        End With

        With colPERIOD8
            .Name = "PERIOD8"
            .HeaderText = "PERIOD 8"
            .Width = 50
        End With

        With colPERIOD9
            .Name = "PERIOD9"
            .HeaderText = "PERIOD 9"
            .Width = 50
        End With

        With colPERIOD10
            .Name = "PERIOD10"
            .HeaderText = "PERIOD 10"
            .Width = 50
        End With

        With colPERIOD11
            .Name = "PERIOD11"
            .HeaderText = "PERIOD 11"
            .Width = 50
        End With

        With colPERIOD12
            .Name = "PERIOD12"
            .HeaderText = "PERIOD 12"
            .Width = 50
        End With

        With colALT_VENDOR
            .Name = "ALT_VENDOR"
            .HeaderText = "ALT VENDOR"
        End With

        With colItemNumber1
            .Name = "ITEMNUMBER1"
            .HeaderText = "ITEM NUMBER - Alpha"
        End With

        With colItemNumber2
            .Name = "ITEMNUMBER2"
            .HeaderText = "ITEM NUMBER - Numeric"
        End With

        With colItemNumberLength
            .Name = "ITEMNUMBERLENGTH"
            .HeaderText = "ITEM NUMBER - Length"
        End With

        With colItemNumberFullLength
            .Name = "ITEMNUMBERFULLLENGTH"
            .HeaderText = "ITEM NUMBER - FullLength"
        End With

        'With colQtyByStore
        '    .Name = "QTYBYSTORE"
        '    .HeaderText = "QTY BY STORE"
        'End With

        With st1Qty
            .Name = "ST1QTY"
            .HeaderText = "ST1QTY"
        End With

        With st2Qty
            .Name = "ST2QTY"
            .HeaderText = "ST2QTY"
        End With

        With st3Qty
            .Name = "ST3QTY"
            .HeaderText = "ST3QTY"
        End With

        With st4Qty
            .Name = "ST4QTY"
            .HeaderText = "ST4QTY"
        End With

        With st6Qty
            .Name = "ST6QTY"
            .HeaderText = "ST6QTY"
        End With

        With st7Qty
            .Name = "ST7QTY"
            .HeaderText = "ST7QTY"
        End With

        With st9Qty
            .Name = "ST9QTY"
            .HeaderText = "ST9QTY"
        End With


        'UO REPORT
        With colSuggestedOrder
            .Name = "SUGGESTED ORDER"
            .HeaderText = "SUGGESTED ORDER"
        End With

        With colYTDUnits
            .Name = "YTDunits"
            .HeaderText = "YTD Units"
        End With

        With colLYRUnits
            .Name = "LYRUnits"
            .HeaderText = "LYR Units"
        End With

        With colAvgCost
            .Name = "AvgCost"
            .HeaderText = "Avg Cost"
        End With

        With colPopCode
            .Name = "PopCode"
            .HeaderText = "Pop Code"
        End With

        With colDateLastSale
            .Name = "DateLastSale"
            .HeaderText = "Date Last Sale"
        End With
        With colDateLastReceipt
            .Name = "DateLastReceipt"
            .HeaderText = "Date Last Receipt"
        End With

        With colTurnsPlus
            .Name = "TurnsPlus"
            .HeaderText = "Turns+"
        End With

        With colStoreCloseout
            .Name = "StoreCloseout"
            .HeaderText = "Store Closeout"
        End With

        With colDiscontinued
            .Name = "Discontinued"
            .HeaderText = "Discontinued"
        End With

        With colCurrSalesUnits
            .Name = "CurrSalesUnits"
            .HeaderText = "Curr Sales Units"
        End With

        With colRetail
            .Name = "Retail"
            .HeaderText = "Retail"
        End With

        With colUnusedS217
            .Name = "UnusedS217"
            .HeaderText = "Unused S217"
        End With


        With dgvData
            If .ColumnCount > 0 Then
                .Columns.Insert(0, strStore)

                If intReportId = 0 Then ' OU
                    .Columns.Insert(4, colQOH)
                    .Columns(4).Width = 50

                    .Columns.Insert(5, colQOO)
                    .Columns(5).Width = 50

                    .Columns.Insert(6, colTOTAL)
                    .Columns(6).Width = 50

                    .Columns.Insert(7, colORDER_POINT)
                    .Columns(7).Width = 50

                    .Columns.Insert(8, colSuggestedOrder)
                    .Columns(8).Width = 50

                    .Columns.Insert(14, colYTDUnits)
                    .Columns(14).Width = 50

                    .Columns.Insert(15, colLYRUnits)
                    .Columns(15).Width = 50

                    .Columns.Insert(16, colAvgCost)
                    .Columns(15).Width = 50

                    .Columns.Insert(17, colPopCode)
                    .Columns(17).Width = 50

                    .Columns.Insert(18, colDateLastSale)
                    .Columns(18).Width = 70

                    .Columns.Insert(19, colDateLastReceipt)
                    .Columns(19).Width = 70

                    .Columns.Insert(20, colTurnsPlus)
                    .Columns(20).Width = 30

                    .Columns.Insert(21, colStoreCloseout)
                    .Columns.Insert(22, colDiscontinued)
                    .Columns.Insert(23, colCurrSalesUnits)

                    .Columns.Insert(24, colPERIOD1)
                    .Columns.Insert(25, colPERIOD2)
                    .Columns.Insert(26, colPERIOD3)
                    .Columns.Insert(27, colPERIOD4)
                    .Columns.Insert(28, colPERIOD5)
                    .Columns.Insert(29, colPERIOD6)
                    .Columns.Insert(30, colPERIOD7)
                    .Columns.Insert(31, colPERIOD8)
                    .Columns.Insert(32, colPERIOD9)
                    .Columns.Insert(33, colPERIOD10)
                    .Columns.Insert(34, colPERIOD11)
                    .Columns.Insert(35, colPERIOD12)

                    .Columns.Insert(36, colRetail)
                    .Columns.Insert(37, colUnusedS217)


                ElseIf intReportId = 2 Then ' ZReport t by vendor
                    .Columns.Insert(6, colQOH)
                    .Columns.Insert(7, colQOO)
                    .Columns.Insert(8, colQTY_AVAIL)
                    .Columns.Insert(9, colORDER_POINT)
                    .Columns.Insert(10, colTOTAL)
                    .Columns.Insert(11, colPERIOD1)
                    .Columns.Insert(12, colPERIOD2)
                    .Columns.Insert(13, colPERIOD3)
                    .Columns.Insert(14, colPERIOD4)
                    .Columns.Insert(15, colPERIOD5)
                    .Columns.Insert(16, colPERIOD6)
                    .Columns.Insert(17, colPERIOD7)
                    .Columns.Insert(18, colPERIOD8)
                    .Columns.Insert(19, colPERIOD9)
                    .Columns.Insert(20, colPERIOD10)
                    .Columns.Insert(21, colPERIOD11)
                    .Columns.Insert(22, colPERIOD12)
                    .Columns.Insert(29, colALT_VENDOR)
                    .Columns.Insert(.Columns.Count, colItemNumber1)
                    .Columns.Insert(.Columns.Count, colItemNumber2)
                    .Columns.Insert(.Columns.Count, colItemNumberLength)
                    .Columns.Insert(.Columns.Count, colItemNumberFullLength)

                    'Added 9/9/2024
                    '.Columns.Insert(.Columns.Count, colQtyByStore)
                    .Columns.Insert(.Columns.Count, st1Qty)
                    .Columns.Insert(.Columns.Count, st2Qty)
                    .Columns.Insert(.Columns.Count, st3Qty)
                    .Columns.Insert(.Columns.Count, st4Qty)
                    .Columns.Insert(.Columns.Count, st6Qty)
                    .Columns.Insert(.Columns.Count, st7Qty)
                    .Columns.Insert(.Columns.Count, st9Qty)

                End If
            End If
        End With

    End Sub
    Private Function getAlpha(ItemNumber As String, isAlpha As Boolean) As String
        Dim strAlpha As String = ""

        For i = 1 To ItemNumber.Length
            If IsNumeric(Mid(ItemNumber, i, 1)) = isAlpha Then
                strAlpha &= Mid(ItemNumber, i, 1)
            End If
        Next
        Return strAlpha

    End Function

    Private Function getNumeric(ItemNumber As String, isDigit As Boolean) As String
        Dim strNum As String = ""

        For i = 1 To ItemNumber.Length
            If IsNumeric(Mid(ItemNumber, i, 1)) = isDigit Then
                strNum &= Mid(ItemNumber, i, 1)
            End If
        Next
        Return strNum

    End Function

    Private Function CommandText(ReportId As Integer) As String
        Dim strCommandText As String = ""

        Select Case ReportId
            Case 0 ' OverStock/UnderStock

                strCommandText = "SELECT DISTINCT [SKU],[Description],[UPC] ,[Posting_Quantity] ,[Primary_Vendor],[Mfg_Part #] ,[Order Multiple] ,[Std Pack] " &
                                    "FROM  [Sheet1$] " &
                                    "ORDER BY SKU"

            Case 1 ' Z Report BY DEPT

            Case 2 ' Z Report BY VENDOR
                strCommandText = "SELECT DISTINCT [ITEM NUMBER], [ITEM DESCRIPTION]" &
                                            ",[PK],[UM] " &
                                            ",(Select TOP 1 [MFG Part #] FROM [Sheet1$] WHERE [ITEM NUMBER] = A.[ITEM NUMBER] ) As [MFG PART #]  " &
                                             ",(Select MAX([REP _COST]) FROM [Sheet1$] WHERE [Item Number] = A.[Item Number]) As [REP COST] " &
                                             ",[RETAIL _PRICE] As [RETAIL PRICE] " &
                                             ",MAX([Date Of Last Sale]) AS [DATE OF LAST SALE] " &
                                             ",MAX([Date Of Last Receipt]) AS [DATE OF LAST RECEIPT]  " &
                                             ",[DEPT],[VENDOR] " &
                                             ",[CLASS],[CODE B1],MAX(LOC) AS [LOC],[DISC ITEM?],[ATTRIBUTE: CUBE]" &
                                         " FROM [Sheet1$] A  " &
                                         " WHERE [CODE B1] Is Not NULL " &
                                         " GROUP BY [Item Number],[ITEM DESCRIPTION], [PK], [UM],[RETAIL _PRICE], [DEPT], [VENDOR], [Class],[Code B1],[Disc Item?],[Attribute: Cube]" &
                                         " ORDER BY [Item Number] "
        End Select

        Return strCommandText

    End Function

    Private Sub loadExcelFile(filename As String)
        ' load Excel File To DataGridView

        My.Application.DoEvents()
        Me.Cursor = Cursors.WaitCursor
        'Dim strCommandText As String

        Try
            'Check if file is open
            Dim isFileOpen As Boolean = False
            Dim excelApp As Excel.Application = Nothing

            Try
                ' Try to get the active instance of Excel
                excelApp = DirectCast(Marshal.GetActiveObject("Excel.Application"), Excel.Application)

            Catch ex As COMException
                ' Active instance of Excel not found, create a new instance
                excelApp = New Excel.Application()
            End Try

            ' Check if the desired file is open
            For Each wb As Excel.Workbook In excelApp.Workbooks
                If String.Compare(wb.FullName, filename, StringComparison.OrdinalIgnoreCase) = 0 Then
                    isFileOpen = True
                    Exit For
                End If

                If blnCancel = True Then
                    Exit Sub
                End If
            Next

            If isFileOpen = True Then
                MessageBox.Show("File is currently open, please close the file and try again " & vbCrLf & Path.GetFileName(filename), "File", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                pbrProgress.Visible = False
                butProcess.Enabled = False
                Me.Cursor = Cursors.Default
                Exit Sub

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message & vbCrLf & vbCrLf & "ERROR FOUND: Check if Excel file is open.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End

        End Try
        '-----------------------------------------------------------------------------

        xlDataAdapter.TableMappings.Clear()
        xlDataset.Clear()

        'Open Database Connection
        'Try
        xlConn = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source= " & filename & ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'")
        If xlConn.State = 1 Then
            xlConn.Close()
        End If
        My.Application.DoEvents()
        lblRows.Visible = True
        lblRows.Text = "Please wait while connecting to the database..."
        xlConn.Open()
        My.Application.DoEvents()

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message & vbCrLf & vbCrLf & "ERROR FOUND: Open Database Connection", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End
        'End Try

        'Main Query Look Up
        'Try
        xlCmd.Connection = xlConn
        My.Application.DoEvents()
        lblRows.Visible = True
        lblRows.Text = "Please wait while importing report..."

        xlCmd.CommandType = CommandType.Text
        xlCmd.CommandText = CommandText(cboReportType.SelectedIndex)
        My.Application.DoEvents()
        lblRows.Visible = True
        lblRows.Text = "Please wait while executing query..."
        xlCmd.ExecuteNonQuery()

        My.Application.DoEvents()
        lblRows.Visible = True
        'lblRows.Text = "Please wait while executing query..."

        xlDataAdapter.SelectCommand = xlCmd
        xlDataAdapter.TableMappings.Add("Table", "Sheet1")
        My.Application.DoEvents()
        lblRows.Visible = True
        lblRows.Text = "Please wait while setting up a dataset..."
        xlDataAdapter.Fill(xlDataset)
        My.Application.DoEvents()

        'Catch ex As Exception
        '    MessageBox.Show("Invalid Selection. Please Check the Report Type", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    lblRows.Text = ""
        '    Me.Cursor = Cursors.Default
        '    Exit Sub

        'End Try

        'Try
        'MAYBE it doesn't need to show the data in datagridview
        With dgvData
            If .Rows.Count > 0 Then
                .DataSource = Nothing
                .Rows.Clear()
                .Columns.Clear()
            End If

            .DataSource = xlDataset
            .DataMember = "Sheet1"
            lblRows.Visible = True

            If .RowCount > 0 Then
                lblRows.Text = "Total Rows:" & .RowCount

                If intReportId = 2 Then 'z-report by vendor
                    strVendorName = .Rows(0).Cells("VENDOR").Value.ToString
                End If

            End If

            DataGridView_InsertColumns(intReportId)

            'z-report by vendor
            .Columns(0).Width = 40
            If intReportId = 2 Then
                .Columns(2).Width = 210
                .Columns(3).Width = 40
                .Columns(4).Width = 40
                .Columns(6).Width = 40
                .Columns(7).Width = 40
                .Columns(8).Width = 40
                .Columns(9).Width = 40
                .Columns(10).Width = 40
            End If

        End With

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message & vbCrLf & "DataridView.DataSource", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        'End Try

        'pbrProgress.Visible = False
        'butProcess.Enabled = True
        'butRefresh.Enabled = True
        isExcelFileLoaded = True
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub LoadFiles(ByVal FileDirectory As String)

        Me.Cursor = Cursors.WaitCursor
        pbrProgress.Visible = True
        pbrProgress.Style = ProgressBarStyle.Marquee

        Dim dDirectories() As String = IO.Directory.GetDirectories(FileDirectory)
        Dim strDirectory As String = ""

        For Each Dir As String In dDirectories

            My.Application.DoEvents()
            Dim dFiles As New IO.DirectoryInfo(Dir.ToString)
            Dim dGetFiles As IO.FileInfo() = dFiles.GetFiles()
            Dim dFile As IO.FileInfo
            Dim strItem(4) As String
            Dim lsvItem As ListViewItem

            For Each dFile In dGetFiles
                My.Application.DoEvents()
                If LCase(Mid(Path.GetFileNameWithoutExtension(dFile.ToString), 1, 4)) = "expo" Then
                    If Path.GetExtension(dFile.ToString) = ".xlsx" Or Path.GetExtension(dFile.ToString) = ".xls" Then

                        strItem(0) = Path.GetFileName(dFile.ToString)
                        strItem(1) = Dir.ToString

                        lsvItem = New ListViewItem(strItem)
                        lsvFiles.Items.Add(lsvItem)
                        lsvFiles.Sorting = System.Windows.Forms.SortOrder.Ascending
                        lsvFiles.Sort()

                    End If
                End If
            Next

            dFiles = Nothing
            dGetFiles = Nothing
        Next

        pbrProgress.Visible = False
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Cursor = Cursors.WaitCursor

        chkProcessExport.Checked = My.Settings.ProcessExport
        cboReportType.SelectedIndex = My.Settings.ReportId
        cboStore.SelectedIndex = My.Settings.StoreId

        'rbItemNumber.Checked = My.Settings.OrderByItemNumber
        'rbDeptClass.Checked = My.Settings.OrderByDept
        'chkSuggestedOrder.Checked = My.Settings.SuggestedOrder
        'rbSuggestedOrder.Checked = My.Settings.SuggestedOrder
        'rbAverageSales.Checked = My.Settings.AverageSales

        cboSortBy.Text = My.Settings.SortBy
        cboOrderPoint.Text = My.Settings.OrderPoint

        'Check Epicor directory
        If Directory.Exists(DirPathEpicor) = False Then
            Directory.CreateDirectory("C:\Users\" & Environment.UserName & "\AppData\Local\Temp\Epicor")
        End If

        If Directory.Exists(DirPath3Apps) = False Then
            MessageBox.Show("Eagle 3apps file directory doesn't exist. Please contact your administrator", "Directory not found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End
        End If

        LoadFiles(DirPathEpicor)
        LoadFiles(DirPath3Apps)

        lsvFiles.Enabled = True
        Me.Cursor = Cursors.Default

        isFormLoaded = True

    End Sub

    Private Sub lsvFiles_Click(sender As Object, e As EventArgs) Handles lsvFiles.Click

        butProcess.Enabled = False
        butExport.Enabled = False
        butRefresh.Enabled = False
        lblProcessingRow.Visible = False

        Dim intIndex = lsvFiles.FocusedItem.Index
        strExcel_FileName = lsvFiles.Items(intIndex).SubItems(0).Text
        strExcel_Path = lsvFiles.Items(intIndex).SubItems(1).Text

        'bgWorker.RunWorkerAsync()
        loadExcelFile(strExcel_Path & "\" & strExcel_FileName)

        'auto-process and export
        If chkProcessExport.Checked = True Then
            Call Process()
            Call Export()
        End If

        butProcess.Enabled = True
        butExport.Enabled = True
        butRefresh.Enabled = True

    End Sub

    Private Function ValidateItemNumber(strItemNumber As String, strChar As String) As String
        Dim strItemNum As String = ""

        For i = 0 To strItemNumber.Length - 1
            If Mid(strItemNumber, i + 1, 1) = "'" Then
                strItemNum &= Mid(strItemNumber, i + 1, 1) & "'"
            Else
                strItemNum &= Mid(strItemNumber, i + 1, 1)
            End If
        Next

        Return strItemNum

    End Function

    Private Function AverageSales(Row As Integer) As Integer
        Dim intPeriodCount As Integer
        Dim dblPeriodTotal As Double

        '(4mos) Ave. Sales = total of 4 mos / 4
        '(3mos) Ave. Sales = total of 3 mos / 3
        '(2mos) Ave. Sales = total of 2 mos / 2
        '(1 mo) Ave. Sales  = total of 1 mo / 2

        For i = 2 To 12
            If CInt(dgvData.Rows(Row).Cells("PERIOD" & i).Value.ToString) > 0 Then
                intPeriodCount += 1
                dblPeriodTotal += CInt(dgvData.Rows(Row).Cells("PERIOD" & i).Value.ToString)
                If intPeriodCount = 4 Then
                    Exit For
                End If
            End If
        Next

        If intPeriodCount = 4 Then
            dblPeriodTotal = dblPeriodTotal / 4

        ElseIf intPeriodCount = 3 Then
            dblPeriodTotal = dblPeriodTotal / 3

        ElseIf intPeriodCount <= 2 Then
            dblPeriodTotal = dblPeriodTotal / 2

        End If

        dblPeriodTotal = Math.Ceiling(dblPeriodTotal)

        If dblPeriodTotal < 0 Then
            dblPeriodTotal = 0
        End If

        Return CInt(dblPeriodTotal)

    End Function

    Private Function AverageSalesBy6Mo(Row As Integer) As Integer
        Dim intPeriodCount As Integer
        Dim dblPeriodTotal As Double

        'Average Sales = (total sales of previous 6 months (not including current month)  / 6
        '(5mos) Ave. Sales = total of 5 mos / 5
        '(4mos) Ave. Sales = total of 4 mos / 4
        '(3mos) Ave. Sales = total of 3 mos / 3
        '(2mos) Ave. Sales = total of 2 mos / 2
        '(1 mo) Ave. Sales  = total of 1 mo / 2

        For i = 2 To 12
            If CInt(dgvData.Rows(Row).Cells("PERIOD" & i).Value.ToString) > 0 Then
                intPeriodCount += 1
                dblPeriodTotal += CInt(dgvData.Rows(Row).Cells("PERIOD" & i).Value.ToString)
                If intPeriodCount = 6 Then
                    Exit For
                End If
            End If
        Next

        Select Case intPeriodCount
            Case 6
                dblPeriodTotal = dblPeriodTotal / 6
            Case 5
                dblPeriodTotal = dblPeriodTotal / 5
            Case 4
                dblPeriodTotal = dblPeriodTotal / 4
            Case 3
                dblPeriodTotal = dblPeriodTotal / 3
            Case <= 2
                dblPeriodTotal = dblPeriodTotal / 2
        End Select

        dblPeriodTotal = Math.Ceiling(dblPeriodTotal)

        If dblPeriodTotal < 0 Then
            dblPeriodTotal = 0
        End If

        Return CInt(dblPeriodTotal)

    End Function

    Private Function SuggestedOrder(Row As Integer) As Double
        Dim intPeriodCount As Integer
        Dim dblPeriodTotal As Double
        Dim dblSuggestedOrder As Double

        '(4mos) Sug. Order = (((total of 4 mos - Qty avail.) / Pack
        '(3mos) Sug. Order = (((total of 3 mos / 3) x 4) - Qty avail.) / Pack
        '(2mos) Sug. Order = (((total of 2 mos / 2) x 4) - Qty avail.) / Pack
        '(1mo)  Sug. Order = (((total of 1 mos / 2) x 4) - Qty avail.) / Pack

        For i = 2 To 12
            If CInt(dgvData.Rows(Row).Cells("PERIOD" & i).Value.ToString) > 0 Then
                intPeriodCount += 1
                dblPeriodTotal += CInt(dgvData.Rows(Row).Cells("PERIOD" & i).Value.ToString)
                If intPeriodCount = 4 Then
                    Exit For
                End If
            End If
        Next

        If intPeriodCount = 4 Then
            dblSuggestedOrder = CInt(dblPeriodTotal)

        ElseIf intPeriodCount = 3 Then
            dblSuggestedOrder = (dblPeriodTotal / 3) * 4

        ElseIf intPeriodCount <= 2 Then
            dblSuggestedOrder = (dblPeriodTotal / 2) * 4

        End If
        dblSuggestedOrder -= CInt(dgvData.Rows(Row).Cells("QTY_AVAIL").Value.ToString)
        dblSuggestedOrder /= CInt(dgvData.Rows(Row).Cells("PK").Value.ToString)
        dblSuggestedOrder = Math.Ceiling(dblSuggestedOrder)

        If dblSuggestedOrder < 0 Then
            dblSuggestedOrder = 0
        End If

        Return dblSuggestedOrder

    End Function

    Private Function SuggestedOrderBy6Mo(Row As Integer) As Double
        Dim intPeriodCount As Integer
        Dim dblPeriodTotal As Double
        Dim dblSuggestedOrder As Double

        'Suggested order = (total sales of previous 6 mos (not including current month) - Qty. available) / Pack
        '(5mos) Sug. Order = (total of 5 mos / 5) x 6 - Qty avail. / Pack
        '(4mos) Sug. Order = (total of 4 mos / 4) x 6 - Qty avail. / Pack
        '(3mos) Sug. Order = (total of 3 mos / 3) x 6 - Qty avail. / Pack
        '(2mos) Sug. Order = (total of 2 mos / 2) x 6 - Qty avail. / Pack
        '(1mo)  Sug. Order = (total of 1 mo / 2) x 6 - Qty avail. / Pack

        For i = 2 To 12
            If CInt(dgvData.Rows(Row).Cells("PERIOD" & i).Value.ToString) > 0 Then
                intPeriodCount += 1
                dblPeriodTotal += CInt(dgvData.Rows(Row).Cells("PERIOD" & i).Value.ToString)
                If intPeriodCount = 6 Then
                    Exit For
                End If
            End If
        Next

        If intPeriodCount > 0 Then

            Select Case intPeriodCount
                Case 6
                    dblSuggestedOrder = dblPeriodTotal
                Case 5
                    dblSuggestedOrder = (dblPeriodTotal / 5) * 6
                Case 4
                    dblSuggestedOrder = (dblPeriodTotal / 4) * 6
                Case 3
                    dblSuggestedOrder = (dblPeriodTotal / 3) * 6
                Case <= 2
                    dblSuggestedOrder = (dblPeriodTotal / 2) * 6
            End Select

            dblSuggestedOrder -= CInt(dgvData.Rows(Row).Cells("QTY_AVAIL").Value.ToString)
            dblSuggestedOrder /= CInt(dgvData.Rows(Row).Cells("PK").Value.ToString)
            dblSuggestedOrder = Math.Ceiling(dblSuggestedOrder)

        End If

        If dblSuggestedOrder < 0 Then
            dblSuggestedOrder = 0
        End If

        Return dblSuggestedOrder

    End Function

    Private Function SuggestedOrderBy12Mo(Row As Integer) As Double
        Dim intPeriodCount As Integer
        Dim dblPeriodTotal As Double
        Dim dblSuggestedOrder As Double

        'Suggested order = (total sales of previous 11 mos (Not including current month) - Qty. available) / Pack
        'Conditions if you found less than 11 months sales:
        '(10mos) Sug. Order = (total of 10 mos / 10) x 11 - Qty avail. / Pack
        '(9mos) Sug. Order = (total of 9 mos / 9) x 11 - Qty avail. / Pack
        '(8mos) Sug. Order = (total of 8 mos / 8) x 11 - Qty avail. / Pack
        '(7mos) Sug. Order = (total of 7 mos / 7) x 11 - Qty avail. / Pack
        '(6mo)  Sug. Order = (total of 6 mo / 6) x 11 - Qty avail. / Pack
        '(5mos) Sug. Order = (total of 5 mos / 5) x 11 - Qty avail. / Pack
        '(4mos) Sug. Order = (total of 4 mos / 4) x 11 - Qty avail. / Pack
        '(3mos) Sug. Order = (total of 3 mos / 3) x 11 - Qty avail. / Pack
        '(2mos) Sug. Order = (total of 2 mos / 2) x 11 - Qty avail. / Pack
        '(1mo)  Sug. Order = (total of 1 mo / 2) x 11 - Qty avail. / Pack

        For i = 2 To 12
            If CInt(dgvData.Rows(Row).Cells("PERIOD" & i).Value.ToString) > 0 Then
                intPeriodCount += 1
                dblPeriodTotal += CInt(dgvData.Rows(Row).Cells("PERIOD" & i).Value.ToString)
                If intPeriodCount = 11 Then
                    Exit For
                End If
            End If
        Next

        If intPeriodCount > 0 Then

            Select Case intPeriodCount
                Case 11
                    dblSuggestedOrder = dblPeriodTotal
                Case 10
                    dblSuggestedOrder = (dblPeriodTotal / 10) * 11
                Case 9
                    dblSuggestedOrder = (dblPeriodTotal / 9) * 11
                Case 8
                    dblSuggestedOrder = (dblPeriodTotal / 8) * 11
                Case 7
                    dblSuggestedOrder = (dblPeriodTotal / 7) * 11
                Case 6
                    dblSuggestedOrder = (dblPeriodTotal / 6) * 11
                Case 5
                    dblSuggestedOrder = (dblPeriodTotal / 5) * 11
                Case 4
                    dblSuggestedOrder = (dblPeriodTotal / 4) * 11
                Case 3
                    dblSuggestedOrder = (dblPeriodTotal / 3) * 11
                Case <= 2
                    dblSuggestedOrder = (dblPeriodTotal / 2) * 11
            End Select

            dblSuggestedOrder -= CInt(dgvData.Rows(Row).Cells("QTY_AVAIL").Value.ToString)
            dblSuggestedOrder /= CInt(dgvData.Rows(Row).Cells("PK").Value.ToString)
            dblSuggestedOrder = Math.Ceiling(dblSuggestedOrder)

        End If

        If dblSuggestedOrder < 0 Then
            dblSuggestedOrder = 0
        End If

        Return dblSuggestedOrder

    End Function



    Private Sub Process()
        butProcess.Enabled = False
        'butExport.Enabled = True
        lblProcessingRow.Visible = True

        Dim DtSet As New System.Data.DataSet

        Dim intRowCount = dgvData.Rows.Count
        Dim intTableRowCount = 0
        Dim intTotal_QOH = 0, intTOTAL_QTY_AVAIL = 0, intQOO = 0, intSuggestedOrder = 0, intTotalPERIOD = 0
        Dim dblOrder_Point As Double
        Dim intPERIOD(0 To 11) As Integer
        Dim strSTORE = ""
        Dim strALT_VENDOR = ""
        Dim intCount_Process = 0
        Dim intPostingQty = 0
        Dim intYTDUnits = 0
        Dim intLYRUnits = 0
        Dim decAvgCost As Decimal = 0
        Dim strDateLastSale As String
        Dim strDateLastReceipt As String
        Dim intCurrentSalesUnits As Integer = 0

        'Dim strQtyByStore As String = "" 'Added 9/9/2024

        'Added 9/10/2024
        Dim st1Qty As String
        Dim st2Qty As String
        Dim st3Qty As String
        Dim st4Qty As String
        Dim st6Qty As String
        Dim st7Qty As String
        Dim st9Qty As String



        Me.Cursor = Cursors.WaitCursor
        pbrProgress.Visible = True
        pbrProgress.Maximum = intRowCount
        pbrProgress.Style = ProgressBarStyle.Continuous

        'Main Loop
        For i = 0 To intRowCount - 1
            My.Application.DoEvents()
            intCount_Process += 1

            pbrProgress.Value = i + 1
            lblProcessingRow.Text = "Row: " + (i + 1).ToString

            xlCmd.Connection = xlConn
            xlCmd.CommandType = CommandType.Text
            If intReportId = 0 Then 'ou-report
                xlCmd.CommandText = "SELECT * " &
                                    "FROM [Sheet1$] " &
                                    "WHERE [SKU] Like '" & ValidateItemNumber(dgvData.Rows(i).Cells(1).Value.ToString, "'") & "'" &
                                    "ORDER BY ST"

            ElseIf intReportId = 1 Then 'z-report by dept
                'No command Text yet

            ElseIf intReportId = 2 Then 'z-report by vendor
                xlCmd.CommandText = "SELECT DISTINCT STORE, QOH, QOO, [Qty Available], " &
                                    "[SALES UNITS PERIOD 1] AS PERIOD1," &
                                    "[SALES UNITS PERIOD 2] AS PERIOD2," &
                                    "[SALES UNITS PERIOD 3] AS PERIOD3," &
                                    "[SALES UNITS PERIOD 4] AS PERIOD4," &
                                    "[SALES UNITS PERIOD 5] AS PERIOD5," &
                                    "[SALES UNITS PERIOD 6] AS PERIOD6," &
                                    "[SALES UNITS PERIOD 7] AS PERIOD7," &
                                    "[SALES UNITS PERIOD 8] AS PERIOD8," &
                                    "[SALES UNITS PERIOD 9] AS PERIOD9," &
                                    "[SALES UNITS PERIOD 10] AS PERIOD10, " &
                                    "[SALES UNITS PERIOD 11] AS PERIOD11," &
                                    "[SALES UNITS PERIOD 12] AS PERIOD12," &
                                    "(Select top 1 [Alt Vendor] from [Sheet1$] where [item number] = A.[item number] ) As [ALT VENDOR] " &
                                "FROM [Sheet1$] A " &
                                "WHERE [ITEM NUMBER] Like '" & ValidateItemNumber(dgvData.Rows(i).Cells(1).Value.ToString, "'") & "'"
            End If
            xlCmd.ExecuteNonQuery()

            xlDataAdapter.SelectCommand = xlCmd
            xlDataAdapter.TableMappings.Add(CStr(i + 1), dgvData.Rows(i).Cells(1).Value.ToString)
            xlDataAdapter.Fill(DtSet)

            'update column header 'months
            If intReportId = 0 Then
                dgvData.Columns(24).HeaderText = DtSet.Tables(0).Columns(25).ColumnName
                dgvData.Columns(25).HeaderText = DtSet.Tables(0).Columns(26).ColumnName
                dgvData.Columns(26).HeaderText = DtSet.Tables(0).Columns(27).ColumnName
                dgvData.Columns(27).HeaderText = DtSet.Tables(0).Columns(28).ColumnName
                dgvData.Columns(28).HeaderText = DtSet.Tables(0).Columns(29).ColumnName
                dgvData.Columns(29).HeaderText = DtSet.Tables(0).Columns(30).ColumnName
                dgvData.Columns(30).HeaderText = DtSet.Tables(0).Columns(31).ColumnName
                dgvData.Columns(31).HeaderText = DtSet.Tables(0).Columns(32).ColumnName
                dgvData.Columns(32).HeaderText = DtSet.Tables(0).Columns(33).ColumnName
                dgvData.Columns(33).HeaderText = DtSet.Tables(0).Columns(34).ColumnName
                dgvData.Columns(34).HeaderText = DtSet.Tables(0).Columns(35).ColumnName
                dgvData.Columns(35).HeaderText = DtSet.Tables(0).Columns(36).ColumnName
            End If


            intTableRowCount = DtSet.Tables(0).Rows.Count

            intTotal_QOH = 0
            intQOO = 0
            intTOTAL_QTY_AVAIL = 0
            dblOrder_Point = 0
            intSuggestedOrder = 0
            intTotalPERIOD = 0
            intPostingQty = 0
            intYTDUnits = 0
            intLYRUnits = 0
            decAvgCost = 0

            strDateLastSale = ""
            strDateLastReceipt = ""
            intCurrentSalesUnits = 0

            'SUB LOOP
            'Added 9/9/2024
            'strQtyByStore = ""
            st1Qty = ""
            st2Qty = ""
            st3Qty = ""
            st4Qty = ""
            st6Qty = ""
            st7Qty = ""
            st9Qty = ""

            'Added 9/10/2024
            For j = 0 To intTableRowCount - 1

                'Added 9/9/2024
                ' For Z-Report only
                'If intReportId = 1 Or intReportId = 2 Then
                '    Dim strQOH As String = DtSet.Tables(0).Rows(j).Item("QOH").ToString
                '    If strQOH = "" Then
                '        strQOH = "0"
                '    ElseIf Convert.ToDouble(strQOH) <= 0 Then
                '        strQOH = "0"
                '    End If
                '    strQtyByStore &= "st" & DtSet.Tables(0).Rows(j).Item("STORE").ToString & ":" & strQOH & "/"
                'End If

                If intReportId = 1 Or intReportId = 2 Then
                    'Dim strQOH As String = DtSet.Tables(0).Rows(j).Item("QOH").ToString
                    'If strQOH <> "" And strQOH <> "0" Then
                    '    strQtyByStore &= "st" & DtSet.Tables(0).Rows(j).Item("STORE").ToString & ":" & strQOH & "/"
                    'End If

                    'Added 9/10/2024
                    Select Case DtSet.Tables(0).Rows(j).Item("STORE").ToString
                        Case "1" : st1Qty = DtSet.Tables(0).Rows(j).Item("QOH").ToString
                        Case "2" : st2Qty = DtSet.Tables(0).Rows(j).Item("QOH").ToString
                        Case "3" : st3Qty = DtSet.Tables(0).Rows(j).Item("QOH").ToString
                        Case "4" : st4Qty = DtSet.Tables(0).Rows(j).Item("QOH").ToString
                        Case "6" : st6Qty = DtSet.Tables(0).Rows(j).Item("QOH").ToString
                        Case "7" : st7Qty = DtSet.Tables(0).Rows(j).Item("QOH").ToString
                        Case "9" : st9Qty = DtSet.Tables(0).Rows(j).Item("QOH").ToString
                    End Select

                End If
                '-----------------------------------------------------------------------------------------------------------------

                If DtSet.Tables(0).Rows(j).Item("QOH").ToString <> "" Then
                    intTotal_QOH += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("QOH"))
                End If

                If DtSet.Tables(0).Rows(j).Item("QOO").ToString <> "" Then
                    intQOO += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("QOO"))
                End If

                'OU 
                If intReportId = 0 Then
                    If DtSet.Tables(0).Rows(j).Item("ORDER POINT").ToString <> "" Then
                        dblOrder_Point += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("ORDER POINT"))
                    End If

                    If DtSet.Tables(0).Rows(j).Item("Suggested Order").ToString <> "" Then
                        intSuggestedOrder += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("Suggested Order"))
                    End If

                    If DtSet.Tables(0).Rows(j).Item("Posting_Quantity").ToString <> "" Then
                        intPostingQty += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("Posting_Quantity"))
                    End If

                    If DtSet.Tables(0).Rows(j).Item("YTD Units").ToString <> "" Then
                        intYTDUnits += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("YTD Units"))
                    End If

                    If DtSet.Tables(0).Rows(j).Item("LYR Units").ToString <> "" Then
                        intLYRUnits += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("LYR Units"))
                    End If

                    If DtSet.Tables(0).Rows(j).Item("Avg Cost").ToString <> "" Then
                        decAvgCost += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("Avg Cost"))
                    End If


                    If DtSet.Tables(0).Rows(j).Item("Date Last Sale").ToString <> "" Or IsDBNull(DtSet.Tables(0).Rows(j).Item("Date Last Sale")) = False Then
                        If strDateLastSale = "" Then
                            strDateLastSale = DtSet.Tables(0).Rows(j).Item("Date Last Sale").ToString
                        Else
                            If Convert.ToDateTime(strDateLastSale) < Convert.ToDateTime(DtSet.Tables(0).Rows(j).Item("Date Last Sale")) Then
                                strDateLastSale = DtSet.Tables(0).Rows(j).Item("Date Last Sale").ToString
                            End If
                        End If
                        strDateLastSale = FormatDateTime(Convert.ToDateTime(strDateLastSale), DateFormat.ShortDate)
                    End If


                    If DtSet.Tables(0).Rows(j).Item("Date Last Recpt").ToString <> "" Or IsDBNull(DtSet.Tables(0).Rows(j).Item("Date Last Recpt")) = False Then
                        If strDateLastReceipt = "" Then
                            strDateLastReceipt = DtSet.Tables(0).Rows(j).Item("Date Last Recpt").ToString
                        Else
                            If Convert.ToDateTime(strDateLastReceipt) < Convert.ToDateTime(DtSet.Tables(0).Rows(j).Item("Date Last Recpt")) Then
                                strDateLastReceipt = DtSet.Tables(0).Rows(j).Item("Date Last Recpt").ToString
                            End If
                        End If
                        strDateLastReceipt = FormatDateTime(Convert.ToDateTime(strDateLastReceipt), DateFormat.ShortDate)
                    End If

                    'Store and Average Cost
                    If DtSet.Tables(0).Rows(j).Item("St").ToString = cboStore.Text And DtSet.Tables(0).Rows(j).Item("AVG Cost").ToString <> "" Then
                        decAvgCost = Convert.ToDecimal(DtSet.Tables(0).Rows(j).Item("AVG Cost"))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(24).ToString <> "" Then
                        intCurrentSalesUnits += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(24).ToString)
                    End If

                    If DtSet.Tables(0).Rows(j).Item(25).ToString <> "" Then
                        intPERIOD(0) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(25).ToString)
                    End If

                    If DtSet.Tables(0).Rows(j).Item(26).ToString <> "" Then
                        intPERIOD(1) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(26))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(27).ToString <> "" Then
                        intPERIOD(2) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(27))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(28).ToString <> "" Then
                        intPERIOD(3) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(28))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(29).ToString <> "" Then
                        intPERIOD(4) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(29))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(30).ToString <> "" Then
                        intPERIOD(5) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(30))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(31).ToString <> "" Then
                        intPERIOD(6) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(31))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(32).ToString <> "" Then
                        intPERIOD(7) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(32))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(33).ToString <> "" Then
                        intPERIOD(8) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(33))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(34).ToString <> "" Then
                        intPERIOD(9) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(34))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(35).ToString <> "" Then
                        intPERIOD(10) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(35))
                    End If

                    If DtSet.Tables(0).Rows(j).Item(36).ToString <> "" Then
                        intPERIOD(11) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item(36))
                    End If

                    'Z-REPORTt by vendor
                ElseIf intReportId = 2 Then
                    intPERIOD(0) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD1"))
                    intPERIOD(1) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD2"))
                    intPERIOD(2) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD3"))
                    intPERIOD(3) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD4"))
                    intPERIOD(4) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD5"))
                    intPERIOD(5) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD6"))
                    intPERIOD(6) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD7"))
                    intPERIOD(7) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD8"))
                    intPERIOD(8) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD9"))
                    intPERIOD(9) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD10"))
                    intPERIOD(10) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD11"))
                    intPERIOD(11) += Convert.ToInt32(DtSet.Tables(0).Rows(j).Item("PERIOD12"))

                End If

            Next ' END OF SUB LOOP

            'GET THE TOTAL OF THE MONTHS
            For k = 0 To 11
                intTotalPERIOD += intPERIOD(k)
            Next

            'QTY AVAILABLE = QOH + QOO EVEN THOUGH QOO HASN'T ARRIVED YET
            intTOTAL_QTY_AVAIL = intTotal_QOH + intQOO

            With dgvData.Rows(i)

                .Cells("PERIOD1").Value = intPERIOD(0)
                .Cells("PERIOD2").Value = intPERIOD(1)
                .Cells("PERIOD3").Value = intPERIOD(2)
                .Cells("PERIOD4").Value = intPERIOD(3)
                .Cells("PERIOD5").Value = intPERIOD(4)
                .Cells("PERIOD6").Value = intPERIOD(5)
                .Cells("PERIOD7").Value = intPERIOD(6)
                .Cells("PERIOD8").Value = intPERIOD(7)
                .Cells("PERIOD9").Value = intPERIOD(8)
                .Cells("PERIOD10").Value = intPERIOD(9)
                .Cells("PERIOD11").Value = intPERIOD(10)
                .Cells("PERIOD12").Value = intPERIOD(11)

                If intReportId = 0 Then 'OU REPORT
                    .Cells("STORE").Value = DtSet.Tables(0).Rows(0).Item("st").ToString
                    .Cells("TOTAL").Value = intTOTAL_QTY_AVAIL 'intTotal_QOH + intQOO   'modified when i was in qc
                    .Cells("SUGGESTED ORDER").Value = intSuggestedOrder
                    .Cells("YTDUnits").Value = intYTDUnits
                    .Cells("LYRUnits").Value = intLYRUnits
                    .Cells("AVGCost").Value = decAvgCost
                    .Cells("PopCode").Value = DtSet.Tables(0).Rows(0).Item("Pop Code").ToString
                    .Cells("DateLastSale").Value = strDateLastSale
                    .Cells("DateLastReceipt").Value = strDateLastReceipt
                    .Cells("TurnsPlus").Value = DtSet.Tables(0).Rows(0).Item("Turns+").ToString
                    .Cells("StoreCloseout").Value = DtSet.Tables(0).Rows(0).Item("Store_Closeout").ToString
                    .Cells("Discontinued").Value = DtSet.Tables(0).Rows(0).Item("Discontinued").ToString
                    .Cells("CurrSalesUnits").Value = intCurrentSalesUnits 'DtSet.Tables(0).Rows(0).Item("Curr Sales Units").ToString
                    .Cells("Retail").Value = DtSet.Tables(0).Rows(0).Item("Retail").ToString
                    .Cells("UnusedS217").Value = DtSet.Tables(0).Rows(0).Item("Unused S217").ToString


                    'ElseIf intReportId = 1 Then

                ElseIf intReportId = 2 Then 'z-report by vendor
                    .Cells("STORE").Value = DtSet.Tables(0).Rows(0).Item("STORE").ToString
                    .Cells("ITEM NUMBER").Value = "'" & .Cells("ITEM NUMBER").Value.ToString

                    'Copy Item Number
                    .Cells("ITEMNUMBER1").Value = getNumeric(.Cells("ITEM NUMBER").Value.ToString, False)
                    .Cells("ITEMNUMBER2").Value = "'" & getNumeric(.Cells("ITEM NUMBER").Value.ToString, True)
                    '.Cells("ITEMNUMBER2").Value = getNumeric(.Cells("ITEM NUMBER").Value.ToString, True)  'We removed the zeros at the front 8/7/2023 4:49 pm
                    .Cells("ITEMNUMBERLENGTH").Value = Len(.Cells("ITEMNUMBER2").Value) - 1
                    '.Cells("ITEMNUMBERFULLLENGTH").Value = (Len(.Cells("ITEMNUMBER1").Value) - 1) + (Len(.Cells("ITEMNUMBER2").Value) - 1)
                    .Cells("ITEMNUMBERFULLLENGTH").Value = (Len(.Cells("ITEM NUMBER").Value) - 1)

                    'Added 9/9/2024
                    'If strQtyByStore <> "" Then
                    '    .Cells("QTYBYSTORE").Value = strQtyByStore.Remove(strQtyByStore.Length - 1, 1)
                    'Else
                    '    .Cells("QTYBYSTORE").Value = "0"
                    'End If

                    'Added 9/10/2024
                    .Cells("ST1QTY").Value = If(st1Qty = "0", "", st1Qty)
                    .Cells("ST2QTY").Value = If(st2Qty = "0", "", st2Qty)
                    .Cells("ST3QTY").Value = If(st3Qty = "0", "", st3Qty)
                    .Cells("ST4QTY").Value = If(st4Qty = "0", "", st4Qty)
                    .Cells("ST6QTY").Value = If(st6Qty = "0", "", st6Qty)
                    .Cells("ST7QTY").Value = If(st7Qty = "0", "", st7Qty)
                    .Cells("ST9QTY").Value = If(st9Qty = "0", "", st9Qty)

                    .Cells("ALT_VENDOR").Value = DtSet.Tables(0).Rows(0).Item("ALT VENDOR")
                    .Cells("QTY_AVAIL").Value = intTOTAL_QTY_AVAIL
                    .Cells("TOTAL").Value = intTotalPERIOD

                    'If cboOrderPoint.SelectedIndex = 0 Then
                    '    dblOrder_Point = AverageSales(i)
                    'ElseIf cboOrderPoint.SelectedIndex = 1 Then
                    '    dblOrder_Point = SuggestedOrder(i)
                    'ElseIf cboOrderPoint.SelectedIndex = 2 Then
                    '    dblOrder_Point = 0
                    'End If

                    Select Case cboOrderPoint.SelectedIndex
                        Case 0
                            dblOrder_Point = AverageSales(i)
                        Case 1
                            dblOrder_Point = AverageSalesBy6Mo(i)
                        Case 2
                            dblOrder_Point = SuggestedOrder(i)
                        Case 3
                            dblOrder_Point = SuggestedOrderBy6Mo(i)
                        Case 4
                            dblOrder_Point = SuggestedOrderBy12Mo(i)
                        Case Else
                            dblOrder_Point = 0
                    End Select
                End If
                .Cells("QOH").Value = intTotal_QOH
                .Cells("QOO").Value = intQOO
                '.Cells("TOTAL").Value = intTotal_QOH + intQOO  'modified when i was in qc

                If intReportId = 2 And cboOrderPoint.Text = "None" Then
                    .Cells("ORDER_POINT").Value = ""
                Else
                    .Cells("ORDER_POINT").Value = dblOrder_Point
                End If

            End With

            'Set intPERIOD value to 0
            For j = 0 To 11
                intPERIOD(j) = 0
            Next
            DtSet.Clear()
        Next

        If intCount_Process = intRowCount Then
            pbrProgress.Visible = False
            butProcess.Enabled = False
            butExport.Enabled = True
            lblProcessingRow.Visible = False
        End If
    End Sub
    Private Sub butProcess_Click(sender As Object, e As EventArgs) Handles butProcess.Click

        Me.Cursor = Cursors.WaitCursor
        Process()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub SetMonthHeaders()
        'Set PERIOD Months
        Dim currentDate As DateTime = DateTime.Now
        Dim intCurrentMonth As Integer = currentDate.Month
        Dim intCol = 11
        Try


            For i = 0 To 11
                dgvData.Columns(intCol).HeaderText = UCase(DateAndTime.MonthName(intCurrentMonth, True))
                intCurrentMonth -= 1
                If intCurrentMonth = 0 Then
                    intCurrentMonth = 12
                End If
                intCol += 1
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbCrLf & " Sub SetMonthHeaders")

        End Try

    End Sub

    Private Sub Export()

        Me.Cursor = Cursors.WaitCursor
        butExport.Enabled = False
        lsvFiles.Enabled = False

        Dim currentTime As DateTime = DateTime.Now

        'Create a New Excel workbook And worksheet
        Dim excel As New Excel.Application
        Dim workbook As Excel.Workbook = excel.Workbooks.Add()
        Dim worksheet As Excel.Worksheet = CType(workbook.ActiveSheet, Excel.Worksheet)

        ' Set the column headers in the first row
        pbrProgress.Maximum = dgvData.ColumnCount
        pbrProgress.Visible = True
        pbrProgress.Style = ProgressBarStyle.Continuous

        If intReportId = 2 Then 'z-reportt by vendor
            SetMonthHeaders()
        End If

        'Header
        Dim intColCount = dgvData.ColumnCount - 1
        For i As Integer = 0 To intColCount
            pbrProgress.Value = i
            My.Application.DoEvents()

            If My.Settings.ReportId = 0 Then
                worksheet.Cells(1, i + 1) = dgvData.Columns(i).HeaderText
            Else
                worksheet.Cells(2, i + 1) = dgvData.Columns(i).HeaderText
            End If
        Next

        'Copy the contents of the DataGridView to the worksheet
        pbrProgress.Maximum = dgvData.RowCount
        pbrProgress.Visible = True
        pbrProgress.Style = ProgressBarStyle.Continuous
        lblRows.Visible = True
        lblRows.Text = "Exporting..."
        lblProcessingRow.Visible = True

        'Datagrid to Excel Worksheet
        Dim intRowCount = dgvData.RowCount - 1
        For i As Integer = 0 To intRowCount

            For j As Integer = 0 To intColCount
                'If Not IsNothing(dgvData.Rows(i).Cells(j).Value.ToString()) Then
                If dgvData.Rows(i).Cells(j).Value.ToString() <> "" Then
                    If intReportId = 0 Then
                        worksheet.Cells(i + 2, j + 1) = dgvData.Rows(i).Cells(j).Value.ToString()
                    Else
                        worksheet.Cells(i + 3, j + 1) = dgvData.Rows(i).Cells(j).Value.ToString()
                    End If
                End If
            Next

            lblProcessingRow.Text = "Row: " & i
            pbrProgress.Value = i + 1
        Next

        'Change Worksheet' header layout
        With worksheet

            If intReportId = 2 Then 'Z-REPORTt by vendor
                With .Range("A1", "C1")
                    .Merge()
                    .HorizontalAlignment = XlHAlign.xlHAlignLeft
                    .Font.Bold = True
                    .Font.Size = 12
                    .EntireColumn.RowHeight = 10
                    .Value = strVendorName & "  " & currentTime.Date 'VENDOR NAME
                End With
            End If

            With .Range("B1", "B1")
                .EntireColumn.RowHeight = 10
            End With

            Dim strRange1, strRange2 As String
            If intReportId = 0 Then 'Overstock
                strRange1 = "A1"
                strRange2 = "AL1"
            Else
                strRange1 = "A2"
                strRange2 = "AT2"
            End If

            With .Range(strRange1, strRange2)
                .Borders.LineStyle = BorderStyle.FixedSingle
                .EntireRow.WrapText = True
                .EntireColumn.RowHeight = 40
                .EntireRow.VerticalAlignment = XlVAlign.xlVAlignCenter
                .Font.Size = 9
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .Interior.Color = RGB(141, 180, 227)

                'Freeze Pane
                .Select()
                excel.ActiveWindow.SplitColumn = 3
                excel.ActiveWindow.SplitRow = 2
                excel.ActiveWindow.FreezePanes = True
            End With

            'Set Worksheet Data Font Size
            If intReportId = 0 Then 'Overstock
                strRange1 = "A2"
                strRange2 = "AL"
            Else
                strRange1 = "A3"
                strRange2 = "AL2"
            End If

            With .Range(strRange1, strRange2 & intRowCount + 3)
                .EntireRow.AutoFit()
                .Font.Size = 9
                '.EntireRow.WrapText = True
            End With

            With .Range("B3", "C" & intRowCount + 3)
                .EntireRow.HorizontalAlignment = XlHAlign.xlHAlignLeft
                .EntireRow.VerticalAlignment = XlVAlign.xlVAlignTop
                '.WrapText = True
            End With

            If intReportId = 2 Then
                With .Range("F3", "F" & intRowCount + 3) 'MFG PART #
                    .Interior.Color = RGB(202, 226, 199)
                    .HorizontalAlignment = XlHAlign.xlHAlignRight
                    .VerticalAlignment = XlVAlign.xlVAlignTop
                End With

                With .Range("I3", "I" & intRowCount + 3) ' QTY AVAILABLE
                    '.EntireRow.WrapText = True
                    .Interior.Color = RGB(254, 226, 227)
                End With

                With .Range("J3", "J" & intRowCount + 3) 'ORDER POINT
                    '.EntireRow.WrapText = True
                    .Interior.Color = RGB(228, 248, 254)
                End With

                With .Range("K3", "K" & intRowCount + 3)  'TOTAL
                    '.EntireRow.WrapText = True
                    .Interior.Color = RGB(245, 236, 154)
                End With

                With .Range("X3", "Y" & intRowCount + 3) 'REP COST, RETAIL PRICE
                    .NumberFormat = "#,##0.00"

                    .HorizontalAlignment = XlHAlign.xlHAlignRight
                    .VerticalAlignment = XlVAlign.xlVAlignTop
                    '.EntireRow.WrapText = True
                End With
            End If

            'Set ColumnWidth
            .Range("A:A").ColumnWidth = 3
            .Range("B:B").ColumnWidth = 5
            .Range("F:F").ColumnWidth = 12
            .Range("I:K").ColumnWidth = 5
            .Range("L:W").ColumnWidth = 4
            .Range("X:Y").ColumnWidth = 5
            .Range("Z:AI").ColumnWidth = 8
            .Range("Z:AA").ColumnWidth = 6
            If intReportId = 2 Then 'z-reportt by vendor
                .Range("Z:AA").NumberFormat = "MM/DD/YY"
            End If

            'Worksheet Name

            If strVendorName <> "" Then
                .Name = strVendorName
            End If

            'LineStyle
            Dim strColLineEnd As String = "AM" ' DEFAULT FOR Z-REPORT
            If intReportId = 0 Then
                strColLineEnd = "AL"
            End If

            With .Range("A2", strColLineEnd & intRowCount + 3)
                .EntireColumn.AutoFit()
                .Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous

                'SORTING
                If intReportId = 2 Then 'z-report by vendor
                    Dim sRange As Excel.Range
                    sRange = .Range("A2", "AT" & intRowCount + 3)
                    sRange.Select()
                    If cboSortBy.SelectedIndex = 0 Then ' rbItemNumber.Checked = True
                        sRange.Sort(Key1:=sRange.Range("AJ1"), Key2:=sRange.Range("AL1"), Key3:=sRange.Range("AK1"), Order1:=XlSortOrder.xlAscending, Orientation:=XlSortOrientation.xlSortColumns)

                    Else 'If rbDeptClass.Checked = True Then
                        sRange.Sort(Key1:=sRange.Range("AB1"), Key2:=sRange.Range("AE1"), Key3:=sRange.Range("B1"), Order1:=XlSortOrder.xlAscending, Orientation:=XlSortOrientation.xlSortColumns)

                    End If

                    ' Hide the columns ITEM NUMBER - Alpha	ITEM NUMBER - Numeric	ITEM NUMBER - Length	ITEM NUMBER - FullLength
                    .Range("AJ2", "AM" & intRowCount + 3).EntireColumn.Hidden = True

                End If

            End With


            'Seeting up page
            With .PageSetup
                .BottomMargin = 22
                .CenterFooter = "&P"
                .CenterHorizontally = True
                .FooterMargin = 11
                '.HeaderMargin = 5
                .LeftMargin = 0
                .Orientation = XlPageOrientation.xlLandscape
                .PaperSize = XlPaperSize.xlPaperLegal
                .PrintArea = "B1:" & "AA" & intRowCount + 3
                .PrintTitleRows = "$2:$2"
                .PrintTitleColumns = "$A" & ":$AA"
                .RightFooter = "&D&T"
                .RightMargin = 0
                .TopMargin = 0
                .Zoom = 90
            End With

        End With

        '--------------------------------Save the workbook ------------------------
        Dim strDateTime As String = ""
        Dim strReportFolder As String = ""
        Dim strPrefix = ""


        Select Case intReportId
            Case 0
                strReportFolder = "\Documents\OU-Report\"
                strPrefix = "OU"

            Case 1 'by dept

            Case 2 'b vendor
                strReportFolder = "\Documents\Z-Report\"
                strPrefix = strVendorName
        End Select

        strDateTime = currentTime.Year.ToString + currentTime.Month.ToString + currentTime.Day.ToString + "_" + currentTime.Hour.ToString + currentTime.Minute.ToString + currentTime.Second.ToString

        If Directory.Exists("C:\Users\" & Environment.UserName & strReportFolder) = False Then
            Directory.CreateDirectory("C:\Users\" & Environment.UserName & strReportFolder)
        End If


        workbook.SaveAs("C:\Users\" & Environment.UserName & strReportFolder & strPrefix & Path.GetFileNameWithoutExtension(strExcel_FileName) & "_" & strDateTime & ".xlsx")
        excel.Visible = True

        lblRows.Visible = False
        lblProcessingRow.Visible = False
        pbrProgress.Visible = False

        butExport.Enabled = True
        lsvFiles.Enabled = True

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub butExport_Click(sender As Object, e As EventArgs) Handles butExport.Click

        Call Export()

    End Sub

    Private Sub butExit_Click(sender As Object, e As EventArgs) Handles butExit.Click
        'If blnCancel = False Then
        '    blnCancel = True
        'Else
        If MessageBox.Show(Me, "Are you sure you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
                Close()
            End If
        'End If

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If MessageBox.Show(Me, "Are you sure you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
            Close()
        End If
    End Sub

    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        frmAbout.ShowDialog()
    End Sub

    Private Sub cboReportType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboReportType.SelectedIndexChanged
        intReportId = cboReportType.SelectedIndex

        If intReportId = 0 Then
            cboStore.Enabled = True
            'grbOrderPoint.Enabled = False
            cboSortBy.Enabled = False
            cboOrderPoint.Enabled = False
        Else
            cboStore.Enabled = False
            'grbOrderPoint.Enabled = True
            cboSortBy.Enabled = True
            cboOrderPoint.Enabled = True
        End If

        'If isFormLoaded = True Then
        '    If isExcelFileLoaded = True And isPopMessage = False Then
        '        MessageBox.Show("Chaanging the report type is still under development " & vbCrLf &
        '                "If you want to change the report type please close and re-open the program.", "Report Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '        isPopMessage = True
        '        cboReportType.SelectedIndex = My.Settings.ReportId
        '    Else

        My.Settings.ReportId = cboReportType.SelectedIndex
        My.Settings.Save()

        '    End If
        'End If
    End Sub

    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStore.SelectedIndexChanged
        intStoreId = cboStore.SelectedIndex

        My.Settings.StoreId = cboStore.SelectedIndex
        My.Settings.Save()

    End Sub

    Private Sub butRefresh_Click(sender As Object, e As EventArgs) Handles butRefresh.Click
        Dim intItemCount As Integer = lsvFiles.Items.Count - 1

        For i = 0 To lsvFiles.Items.Count - 1
            If lsvFiles.Items.Count > 0 Then
                lsvFiles.Items.RemoveAt(lsvFiles.Items.Count - 1)
            End If
        Next

        LoadFiles(DirPathEpicor)
        LoadFiles(DirPath3Apps)

    End Sub

    Private Sub cboReportType_Click(sender As Object, e As EventArgs) Handles cboReportType.Click
        'intReportId = cboReportType.SelectedIndex

        'If intReportId = 0 Then
        '    cboStore.Enabled = True
        'Else
        '    cboStore.Enabled = False
        'End If

        'If isFormLoaded = True Then
        '    If isExcelFileLoaded = True And isPopMessage = False Then
        '        MessageBox.Show("Chaanging the report type is still under development " & vbCrLf &
        '                "If you want to change the report type please close and re-open the program.", "Report Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '        isPopMessage = True
        '        cboReportType.SelectedIndex = My.Settings.ReportId
        '    Else

        '        My.Settings.ReportId = cboReportType.SelectedIndex
        '        My.Settings.Save()

        '    End If
        '    End If
    End Sub

    'Private Sub butMultiSelect_Click(sender As Object, e As EventArgs)
    '    If isMultiSelect = False Then
    '        isMultiSelect = True
    '        butMultiSelect.Text = "Multi-select Off"
    '    Else
    '        isMultiSelect = False
    '        butMultiSelect.Text = "Multi-select On"
    '    End If

    '    lsvFiles.CheckBoxes = isMultiSelect
    'End Sub

    'Private Sub rbDeptClass_CheckedChanged(sender As Object, e As EventArgs)
    '    My.Settings.OrderByDept = rbDeptClass.Checked
    '    My.Settings.Save()
    'End Sub

    'Private Sub rbItemNumber_CheckedChanged(sender As Object, e As EventArgs)
    '    My.Settings.OrderByItemNumber = rbItemNumber.Checked
    '    My.Settings.Save()
    'End Sub

    Private Sub chkProcessExport_CheckedChanged(sender As Object, e As EventArgs) Handles chkProcessExport.CheckedChanged
        My.Settings.ProcessExport = chkProcessExport.Checked
        My.Settings.Save()
    End Sub

    Private Sub lsvFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lsvFiles.SelectedIndexChanged

    End Sub

    Private Sub frmMain_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp

    End Sub

    Private Sub cboSortBy_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSortBy.SelectedIndexChanged
        My.Settings.SortBy = cboSortBy.Text
        My.Settings.Save()
    End Sub

    Private Sub cboOrderPoint_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboOrderPoint.SelectedIndexChanged
        My.Settings.OrderPoint = cboOrderPoint.Text
        My.Settings.Save()
    End Sub

    'Private Sub bgWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bgWorker.DoWork
    '    loadExcelFile(strExcel_Path & "\" & strExcel_FileName)
    'End Sub

    'Private Sub chkDept_CheckedChanged(sender As Object, e As EventArgs)
    '    My.Settings.ChkDept = 0
    '    If chkDept.Checked = True Then
    '        My.Settings.ChkDept = 1
    '    End If
    '    My.Settings.Save()
    'End Sub

    'Private Sub chkClass_CheckedChanged(sender As Object, e As EventArgs)
    'My.Settings.ChkDept = 0
    'If chkDept.Checked = True Then
    '    My.Settings.ChkDept = 1
    'End If
    'My.Settings.Save()
    'End Sub

    'Private Sub chkItemNumber_CheckedChanged(sender As Object, e As EventArgs)
    '    My.Settings.ChkItemNumber = 0
    '    If chkItemNumber.Checked = True Then
    '        My.Settings.ChkItemNumber = 1
    '    End If
    '    My.Settings.Save()
    'End Sub
End Class
