Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Common
Imports System.Windows.Input
Imports System.Data.SqlTypes
Imports System.Runtime.InteropServices

Public Class frmMain

    Dim Thread As System.Threading.Thread

    'Database Connection
    Dim xlConn As New OleDb.OleDbConnection
    Dim xlDataset As New System.Data.DataSet
    Dim xlDataAdapter As New System.Data.OleDb.OleDbDataAdapter
    Dim xlCmd As New System.Data.OleDb.OleDbCommand

    Private DirPathEpicor As String = "C:\Users\" & Environment.UserName & "\AppData\Local\Temp\Epicor"
    Private DirPath3Apps As String = "C:\3apps\Temp"
    Private strExcel_Path As String
    Private strExcel_FileName As String
    Private strVendorName As String

    Private Sub DataGridView_InsertColumns()
        Dim strStore As New DataGridViewTextBoxColumn()
        Dim colQOH As New DataGridViewTextBoxColumn()
        Dim colQOO As New DataGridViewTextBoxColumn()
        Dim colQTY_AVAIL As New DataGridViewTextBoxColumn()
        Dim colORDER_POINT As New DataGridViewTextBoxColumn()
        Dim colTOTAL As New DataGridViewTextBoxColumn()
        Dim colALT_VENDOR As New DataGridViewTextBoxColumn()

        Dim colPeriod1 As New DataGridViewTextBoxColumn()
        Dim colPeriod2 As New DataGridViewTextBoxColumn()
        Dim colPeriod3 As New DataGridViewTextBoxColumn()
        Dim colPeriod4 As New DataGridViewTextBoxColumn()
        Dim colPeriod5 As New DataGridViewTextBoxColumn()
        Dim colPeriod6 As New DataGridViewTextBoxColumn()
        Dim colPeriod7 As New DataGridViewTextBoxColumn()
        Dim colPeriod8 As New DataGridViewTextBoxColumn()
        Dim colPeriod9 As New DataGridViewTextBoxColumn()
        Dim colPeriod10 As New DataGridViewTextBoxColumn()
        Dim colPeriod11 As New DataGridViewTextBoxColumn()
        Dim colPeriod12 As New DataGridViewTextBoxColumn()


        strStore.Name = "STORE"
        strStore.HeaderText = "STORE"

        colQOH.Name = "QOH"
        colQOH.HeaderText = "QOH"

        colQOO.Name = "QOO"
        colQOO.HeaderText = "QOO"

        colQTY_AVAIL.Name = "QTY_AVAIL"
        colQTY_AVAIL.HeaderText = "QTY AVAILABLE"

        colORDER_POINT.Name = "ORDER_POINT"
        colORDER_POINT.HeaderText = "ORDER POINT"

        colTOTAL.Name = "TOTAL"
        colTOTAL.HeaderText = "TOTAL"

        colPeriod1.Name = "Period1"
        colPeriod1.HeaderText = "Period 1"
        colPeriod1.Width = 50

        colPeriod2.Name = "Period2"
        colPeriod2.HeaderText = "Period 2"
        colPeriod2.Width = 50

        colPeriod3.Name = "Period3"
        colPeriod3.HeaderText = "Period 3"
        colPeriod3.Width = 50

        colPeriod4.Name = "Period4"
        colPeriod4.HeaderText = "Period 4"
        colPeriod4.Width = 50

        colPeriod5.Name = "Period5"
        colPeriod5.HeaderText = "Period 5"
        colPeriod5.Width = 50

        colPeriod6.Name = "Period6"
        colPeriod6.HeaderText = "Period 6"
        colPeriod6.Width = 50

        colPeriod7.Name = "Period7"
        colPeriod7.HeaderText = "Period 7"
        colPeriod7.Width = 50

        colPeriod8.Name = "Period8"
        colPeriod8.HeaderText = "Period 8"
        colPeriod8.Width = 50

        colPeriod9.Name = "Period9"
        colPeriod9.HeaderText = "Period 9"
        colPeriod9.Width = 50

        colPeriod10.Name = "Period10"
        colPeriod10.HeaderText = "Period 10"
        colPeriod10.Width = 50

        colPeriod11.Name = "Period11"
        colPeriod11.HeaderText = "Period 11"
        colPeriod11.Width = 50

        colPeriod12.Name = "Period12"
        colPeriod12.HeaderText = "Period 12"
        colPeriod12.Width = 50

        colALT_VENDOR.Name = "ALT_VENDOR"
        colALT_VENDOR.HeaderText = "ALT VENDOR"

        With dgvData
            If .ColumnCount > 0 Then
                .Columns.Insert(0, strStore)
                .Columns.Insert(6, colQOH)
                .Columns.Insert(7, colQOO)
                .Columns.Insert(8, colQTY_AVAIL)
                .Columns.Insert(9, colORDER_POINT)
                .Columns.Insert(10, colTOTAL)

                .Columns.Insert(11, colPeriod1)
                .Columns.Insert(12, colPeriod2)
                .Columns.Insert(13, colPeriod3)
                .Columns.Insert(14, colPeriod4)
                .Columns.Insert(15, colPeriod5)
                .Columns.Insert(16, colPeriod6)
                .Columns.Insert(17, colPeriod7)
                .Columns.Insert(18, colPeriod8)
                .Columns.Insert(19, colPeriod9)
                .Columns.Insert(20, colPeriod10)
                .Columns.Insert(21, colPeriod11)
                .Columns.Insert(22, colPeriod12)

                .Columns.Insert(29, colALT_VENDOR)
            End If
        End With

    End Sub

    Private Sub loadExcelFile(filename As String)
        ' load Excel File To DataGridView

        My.Application.DoEvents()
        Me.Cursor = Cursors.WaitCursor


        'Check if file is open
        Dim excelApp As Excel.Application = Nothing
        Dim isFileOpen As Boolean = False

        excelApp = DirectCast(Marshal.GetActiveObject("Excel.Application"), Excel.Application)

        ' Check if the desired file is open
        For Each wb As Excel.Workbook In excelApp.Workbooks
            If String.Compare(wb.FullName, filename, StringComparison.OrdinalIgnoreCase) = 0 Then
                isFileOpen = True
                Exit For
            End If
        Next

        If isFileOpen = True Then
            MessageBox.Show("File is open, close the file and try again " & vbCrLf & Path.GetFileName(filename), "File Open", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            pbrProgress.Visible = False
            butProcess.Enabled = False
            Me.Cursor = Cursors.Default
            Exit Sub

        End If
        '-----------------------------------------------------------------------------

        'pbrProgress.Visible = True
        'pbrProgress.Style = ProgressBarStyle.Marquee

        xlDataAdapter.TableMappings.Clear()
        xlDataset.Clear()

        'Open Database Connection
        xlConn = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source= " & filename & ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'")
        If xlConn.State = 1 Then
            xlConn.Close()
        End If
        My.Application.DoEvents()
        lblRows.Visible = True
        lblRows.Text = "Please wait while connecting to the database..."
        xlConn.Open()
        My.Application.DoEvents()

        'Main Query Look Up
        xlCmd.Connection = xlConn

        My.Application.DoEvents()
        lblRows.Visible = True
        lblRows.Text = "Please wait while importing your report..."

        xlCmd.CommandType = CommandType.Text
        xlCmd.CommandText = "SELECT DISTINCT [ITEM NUMBER], [ITEM DESCRIPTION]" &
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


        'My.Application.DoEvents()
        'lblRows.Visible = True
        'lblRows.Text = "Please wait while executing query..."
        xlCmd.ExecuteNonQuery()
        My.Application.DoEvents()
        xlDataAdapter.SelectCommand = xlCmd
        xlDataAdapter.TableMappings.Add("Table", "Sheet1")

        'My.Application.DoEvents()
        'lblRows.Visible = True
        'lblRows.Text = "Please wait while setting a dataset..."

        xlDataAdapter.Fill(xlDataset)
        My.Application.DoEvents()

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
                strVendorName = .Rows(0).Cells("VENDOR").Value.ToString
            End If

            If .ColumnCount > 0 Then

                DataGridView_InsertColumns()

                .Columns(0).Width = 40
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

        pbrProgress.Visible = False
        butProcess.Enabled = True
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

                If Mid(dFile.ToString, 1, 2) <> "~$" And (Path.GetExtension(dFile.ToString) = ".xlsx" Or Path.GetExtension(dFile.ToString) = ".xls" Or Path.GetExtension(dFile.ToString) = ".csv") Then

                    strItem(0) = Path.GetFileName(dFile.ToString)
                    strItem(1) = Dir.ToString

                    lsvItem = New ListViewItem(strItem)
                    lsvFiles.Items.Add(lsvItem)
                    lsvFiles.Sorting = System.Windows.Forms.SortOrder.Ascending
                    lsvFiles.Sort()

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

    End Sub

    Private Sub lsvFiles_Click(sender As Object, e As EventArgs) Handles lsvFiles.Click

        butProcess.Enabled = False
        butExport.Enabled = False
        lblProcessingRow.Visible = False

        Dim intIndex = lsvFiles.FocusedItem.Index
        strExcel_FileName = lsvFiles.Items(intIndex).SubItems(0).Text
        strExcel_Path = lsvFiles.Items(intIndex).SubItems(1).Text

        loadExcelFile(strExcel_Path & "\" & strExcel_FileName)

        'lblRows.Visible = True
        'butProcess.Enabled = True

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
    Private Sub butProcess_Click(sender As Object, e As EventArgs) Handles butProcess.Click

        Me.Cursor = Cursors.WaitCursor

        butProcess.Enabled = False
        'butExport.Enabled = True
        lblProcessingRow.Visible = True

        Dim Dtset As New System.Data.DataSet

        Dim intRowCount = dgvData.Rows.Count
        Dim intTableRowCount = 0
        Dim intTotal_QOH = 0, intTOTAL_QTY_AVAIL = 0, intQOO = 0, intOrder_Point = 0, intTotalPeriod = 0
        Dim intPeriod(0 To 11) As Integer
        Dim strSTORE = ""
        Dim strALT_VENDOR = ""
        Dim intCount_Process = 0

        Me.Cursor = Cursors.WaitCursor
        pbrProgress.Visible = True
        pbrProgress.Maximum = intRowCount
        pbrProgress.Style = ProgressBarStyle.Continuous

        For i = 0 To intRowCount - 1
            My.Application.DoEvents()
            intCount_Process += 1
            pbrProgress.Value = i + 1
            lblProcessingRow.Text = "Row: " + (i + 1).ToString

            xlCmd.Connection = xlConn
            xlCmd.CommandType = CommandType.Text
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
            xlCmd.ExecuteNonQuery()

            xlDataAdapter.SelectCommand = xlCmd
            xlDataAdapter.TableMappings.Add(CStr(i + 1), dgvData.Rows(i).Cells(1).Value.ToString)
            xlDataAdapter.Fill(Dtset)


            intTableRowCount = Dtset.Tables(0).Rows.Count

            intTotal_QOH = 0
            intQOO = 0
            intTOTAL_QTY_AVAIL = 0
            strSTORE = Dtset.Tables(0).Rows(0).Item("STORE").ToString
            intTotalPeriod = 0

            For j = 0 To intTableRowCount - 1
                If Dtset.Tables(0).Rows(j).Item("QOH").ToString <> "" Then
                    intTotal_QOH += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("QOH"))
                End If

                If Dtset.Tables(0).Rows(0).Item("QOO").ToString <> "" Then
                    intQOO += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("QOO"))
                End If

                'If Dtset.Tables(0).Rows(j).Item("QTY AVAILABLE").ToString <> "" Then
                '    intTOTAL_QTY_AVAIL += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("QTY AVAILABLE"))
                'End If

                intPeriod(0) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD1"))
                intPeriod(1) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD2"))
                intPeriod(2) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD3"))
                intPeriod(3) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD4"))
                intPeriod(4) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD5"))
                intPeriod(5) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD6"))
                intPeriod(6) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD7"))
                intPeriod(7) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD8"))
                intPeriod(8) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD9"))
                intPeriod(9) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD10"))
                intPeriod(10) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD11"))
                intPeriod(11) += Convert.ToInt32(Dtset.Tables(0).Rows(j).Item("PERIOD12"))

            Next

            'GET THE TOTAL OF THE MONTHS
            For k = 0 To 11
                intTotalPeriod += intPeriod(k)
            Next

            'QTY AVAILABLE = QOH + QOO EVEN THOUGH QOO HASN'T ARRIVED YET
            intTOTAL_QTY_AVAIL = intTotal_QOH + intQOO

            'If i < intRowCount Then
            '    dgvData.FirstDisplayedScrollingRowIndex = i

            'End If

            With dgvData.Rows(i)
                .Cells("STORE").Value = strSTORE
                .Cells("ITEM NUMBER").Value = "'" & .Cells("ITEM NUMBER").Value.ToString
                .Cells("QOH").Value = intTotal_QOH
                .Cells("QOO").Value = intQOO
                .Cells("ORDER_POINT").Value = intOrder_Point
                .Cells("QTY_AVAIL").Value = intTOTAL_QTY_AVAIL

                .Cells("ALT_VENDOR").Value = Dtset.Tables(0).Rows(0).Item("ALT VENDOR")

                .Cells("Period1").Value = intPeriod(0)
                .Cells("Period2").Value = intPeriod(1)
                .Cells("Period3").Value = intPeriod(2)
                .Cells("Period4").Value = intPeriod(3)
                .Cells("Period5").Value = intPeriod(4)
                .Cells("Period6").Value = intPeriod(5)
                .Cells("Period7").Value = intPeriod(6)
                .Cells("Period8").Value = intPeriod(7)
                .Cells("Period9").Value = intPeriod(8)
                .Cells("Period10").Value = intPeriod(9)
                .Cells("Period11").Value = intPeriod(10)
                .Cells("Period12").Value = intPeriod(11)

                .Cells("TOTAL").Value = intTotalPeriod  'Convert.ToInt16(.Cells(11).Value.ToString) '+ Convert.ToInt16(.Cells(12).Value.ToString) + Convert.ToInt16(.Cells(13).Value.ToString) + Convert.ToInt16(.Cells(14).Value.ToString) +
                'Convert.ToInt16(.Cells(15).Value.ToString) + Convert.ToInt16(.Cells(16).Value.ToString) + Convert.ToInt16(.Cells(17).Value.ToString) + Convert.ToInt16(.Cells(18).Value.ToString) +
                'Convert.ToInt16(.Cells(19).Value.ToString) + Convert.ToInt16(.Cells(20).Value.ToString) + Convert.ToInt16(.Cells(21).Value.ToString) + Convert.ToInt16(.Cells(22).Value.ToString)

            End With

            'Set intPeriod value to 0
            For j = 0 To 11
                intPeriod(j) = 0
            Next

            Dtset.Clear()
        Next

        If intCount_Process = intRowCount Then
            pbrProgress.Visible = False
            butProcess.Enabled = False
            butExport.Enabled = True
            lblProcessingRow.Visible = False
        End If

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub SetMonthHeaders()
        'Set Period Months
        Dim currentDate As DateTime = DateTime.Now
        Dim intCurrentMonth As Integer = currentDate.Month
        Dim intCol = 11

        For i = 0 To 11
            dgvData.Columns(intCol).HeaderText = UCase(DateAndTime.MonthName(intCurrentMonth, True))
            intCurrentMonth -= 1
            If intCurrentMonth = 0 Then
                intCurrentMonth = 12
            End If
            intCol += 1
        Next
    End Sub

    Private Sub butExport_Click(sender As Object, e As EventArgs) Handles butExport.Click

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

        SetMonthHeaders()

        'VENDOR Name
        With worksheet
            With .Range("A1", "C1")
                .Merge()
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .Font.Bold = True
                .Font.Size = 12
                .EntireColumn.RowHeight = 10
                .Value = strVendorName & "  " & currentTime.Date
            End With

            'Date File Created
            'With .Range("D1", "D1")
            '    .Value = "Date" & File.GetCreationTime(strExcel_Path & "\" & strExcel_FileName)
            '    .WrapText = True
            'End With
        End With

        'Header
        Dim intColCount = dgvData.ColumnCount - 1
        For i As Integer = 0 To intColCount
            pbrProgress.Value = i
            My.Application.DoEvents()
            worksheet.Cells(2, i + 1) = dgvData.Columns(i).HeaderText
        Next

        'Copy the contents of the DataGridView to the worksheet
        pbrProgress.Maximum = dgvData.RowCount
        pbrProgress.Visible = True
        pbrProgress.Style = ProgressBarStyle.Continuous
        lblRows.Visible = True
        lblRows.Text = "Exporting..."
        lblProcessingRow.Visible = True

        Dim intRowCount = dgvData.RowCount - 1
        For i As Integer = 0 To intRowCount

            For j As Integer = 0 To intColCount
                If Not IsNothing(dgvData.Rows(i).Cells(j).Value.ToString()) Then
                    worksheet.Cells(i + 3, j + 1) = dgvData.Rows(i).Cells(j).Value.ToString()
                End If
            Next

            lblProcessingRow.Text = "Row: " & i
            pbrProgress.Value = i + 1
        Next

        'Change Worksheet' header layout
        With worksheet

            With .Range("A1", "C1")
                .Merge()
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .Font.Bold = True
                .Font.Size = 12
                .EntireColumn.RowHeight = 10
                .Value = strVendorName & "  " & currentTime.Date 'VENDOR NAME
            End With


            With .Range("B1", "B1")
                .EntireColumn.RowHeight = 10
                '.EntireRow.AutoFit()
                '.EntireRow.WrapText = True
            End With

            With .Range("A2", "AI2")
                .Borders.LineStyle = BorderStyle.FixedSingle
                '.EntireColumn.AutoFit()
                '.EntireRow.AutoFit()
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

            'Set Worksheet Font Size
            With .Range("A3", "AI" & intRowCount + 3)
                .EntireRow.AutoFit()
                .Font.Size = 9
                '.EntireRow.WrapText = True
            End With

            With .Range("B3", "C" & intRowCount + 3)
                .EntireRow.HorizontalAlignment = XlHAlign.xlHAlignLeft
                .EntireRow.VerticalAlignment = XlVAlign.xlVAlignTop
                '.WrapText = True
            End With

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


            'Set ColumnWidth
            .Range("A:A").ColumnWidth = 3
            .Range("B:B").ColumnWidth = 5
            .Range("F:F").ColumnWidth = 12
            .Range("I:K").ColumnWidth = 5
            .Range("L:W").ColumnWidth = 4
            .Range("X:Y").ColumnWidth = 5
            .Range("Z:AI").ColumnWidth = 8
            .Range("Z:AA").ColumnWidth = 6
            .Range("Z:AA").NumberFormat = "M/DD/YY"

            'Worksheet Name
            .Name = strVendorName

            'LineStyle
            With .Range("A2", "AI" & intRowCount + 3)
                .EntireColumn.AutoFit()
                .Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                .Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
            End With

            'With .Range("B1")
            '.EntireColumn.RowHeight = 10
            '.EntireRow.AutoFit()
            '.EntireRow.WrapText = True
            'End With

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
        Dim strDateTime As String

        strDateTime = currentTime.Year.ToString + currentTime.Month.ToString + currentTime.Day.ToString + "_" + currentTime.Hour.ToString + currentTime.Minute.ToString + currentTime.Second.ToString

        If Directory.Exists("C:\Users\" & Environment.UserName & "\Documents\Z-Report\") = False Then
            Directory.CreateDirectory("C:\Users\" & Environment.UserName & "\Documents\Z-Report\")
        End If

        workbook.SaveAs("C:\Users\" & Environment.UserName & "\Documents\Z-Report\" & strVendorName & "_" & Path.GetFileNameWithoutExtension(strExcel_FileName) & "_" & strDateTime & ".xlsx")
        excel.Visible = True

        lblRows.Visible = False
        lblProcessingRow.Visible = False
        pbrProgress.Visible = False

        butExport.Enabled = True
        lsvFiles.Enabled = True

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub butExit_Click(sender As Object, e As EventArgs) Handles butExit.Click
        If MessageBox.Show(Me, "Are you sure you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
            Close()
        End If
    End Sub

    Private Sub frmMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '    If MessageBox.Show(Me, "Are you sure you want to exit?", "Exit?", MessageBoxButtons.YesNo) = vbYes Then
        '        Close()
        '    End If
    End Sub

    Private Sub lsvFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lsvFiles.SelectedIndexChanged

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If MessageBox.Show(Me, "Are you sure you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
            Close()
        End If
    End Sub

    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        frmAbout.ShowDialog()
    End Sub
End Class
