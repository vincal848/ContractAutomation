' SUPERSEDED. Kept for reference only -- this file is not part of the importable
' package, does not compile, and is known to be incorrect. Replaced by
' src/ReportAutomation.bas.
'
' Defects in this file (see also Financial_report_module.bas, the earlier draft
' this one improves on but does not fully fix):
'   1. TransferToWordTemplate and AutomateFinancialReporting share their exact
'      names with the subs in Financial_report_module.bas. Importing both
'      modules into one workbook raises "ambiguous name detected".
'   2. AutomateFinancialReporting uses "Continue Do" to skip to the next
'      DataSources row. Continue Do/For/While does not exist in VBA (it is a
'      VB.NET/C# construct), so this module fails to compile.
'   3. TransferToWordTemplate opens one wordDoc before the loop and reuses it
'      for every row. A bookmark's Range collapses to its filled text the
'      first time it is written, so only the first row's contract is actually
'      filled in; every later row finds the bookmarks already consumed and
'      SaveAs2 just writes copies of the first row's document under different
'      file names.
'   4. AutomateFinancialReporting reads the next DataSources path with
'      wsSources.Cells(currentRow - 1, 1), where currentRow is still derived
'      from RawData's row count rather than a counter over DataSources itself.
'      The off-by-one offset does not fix the underlying desync described in
'      Financial_report_module.bas defect 4; it only changes where it shows up.
'   5. FormatSummaryReport calls ListObjects.Add on every run. A second run
'      against a workbook that already has the table fails with "a table
'      already exists on this worksheet".
'   6. CalculateMetrics divides by wsData.Cells(i - 1, 3).Value for revenue
'      growth with no check for zero, so a month with zero prior revenue
'      raises a division-by-zero error.
'   7. CreateCharts calls SetSourceData on Range("A1:B" & lastRow) without
'      pinning PlotBy, so Excel's own row-vs-column guess is free to mislabel
'      the series: the sample screenshot shows the first data row's values
'      used as legend text ("Net Profit $40,000.00", "Revenue Growth (%) N/A")
'      instead of the header row.
'   8. TransferToWordTemplate reads the template path from TemplateInfo!B1 (a
'      real fix over Financial_report_module.bas) but still hardcodes the
'      output folder as "C:\Contracts\" and never creates it if missing.
'
Attribute VB_Name = "Module2"
Sub TransferToWordTemplate()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim filePath As String
    
    ' Get the template path from the TemplateInfo sheet
    Dim templateSheet As Worksheet
    On Error GoTo ErrorHandler
    Set templateSheet = ThisWorkbook.Sheets("TemplateInfo")
    filePath = templateSheet.Range("B1").Value
    
    ' Check if the file path is valid
    If Dir(filePath) = "" Then
        MsgBox "The Word template file could not be found. Please check the path in TemplateInfo sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Set up Word application
    On Error Resume Next
    Set wordApp = CreateObject("Word.Application")
    On Error GoTo ErrorHandler
    
    ' If Word is not running, create a new instance
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    
    ' Open the Word template
    Set wordDoc = wordApp.Documents.Open(filePath)
    
    ' Optional: make Word visible
    wordApp.Visible = True

    ' Assuming your Excel data is in RawData
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets("RawData")
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    ' Loop through each region data to create contracts
    Dim i As Long
    For i = 2 To lastRow
        ' Ensure data exists in all required columns
        If IsEmpty(wsData.Cells(i, 1)) Or IsEmpty(wsData.Cells(i, 2)) Or _
           IsEmpty(wsData.Cells(i, 3)) Or IsEmpty(wsData.Cells(i, 4)) Or _
           IsEmpty(wsData.Cells(i, 5)) Or IsEmpty(wsData.Cells(i, 6)) Then
            MsgBox "Missing data in row " & i & ". Please verify the RawData sheet.", vbExclamation
            Exit Sub
        End If

        ' Set up placeholders in the Word document with regional data
        Dim region As String, month As String, revenue As Double, expenses As Double, netProfit As Double
        Dim customerName As String, companyName As String
        
        region = wsData.Cells(i, 1).Value
        month = wsData.Cells(i, 2).Value
        revenue = wsData.Cells(i, 3).Value
        expenses = wsData.Cells(i, 4).Value
        netProfit = revenue - expenses
        customerName = wsData.Cells(i, 5).Value
        companyName = wsData.Cells(i, 6).Value
        
        ' Ensure wordDoc is valid before trying to access bookmarks
        If Not wordDoc Is Nothing Then
            ' Check if bookmarks exist before trying to fill them
            If wordDoc.Bookmarks.Exists("RegionBookmark") Then wordDoc.Bookmarks("RegionBookmark").Range.Text = region
            If wordDoc.Bookmarks.Exists("MonthBookmark") Then wordDoc.Bookmarks("MonthBookmark").Range.Text = month
            If wordDoc.Bookmarks.Exists("RevenueBookmark") Then wordDoc.Bookmarks("RevenueBookmark").Range.Text = Format(revenue, "$#,##0.00")
            If wordDoc.Bookmarks.Exists("ExpensesBookmark") Then wordDoc.Bookmarks("ExpensesBookmark").Range.Text = Format(expenses, "$#,##0.00")
            If wordDoc.Bookmarks.Exists("NetProfitBookmark") Then wordDoc.Bookmarks("NetProfitBookmark").Range.Text = Format(netProfit, "$#,##0.00")
            If wordDoc.Bookmarks.Exists("CustomerNameBookmark") Then wordDoc.Bookmarks("CustomerNameBookmark").Range.Text = customerName
            If wordDoc.Bookmarks.Exists("CompanyNameBookmark") Then wordDoc.Bookmarks("CompanyNameBookmark").Range.Text = companyName
            
            ' Save the filled contract with a unique name
            Dim outputPath As String
            outputPath = "C:\Contracts\" & region & "_" & month & "_Contract.docx"
            wordDoc.SaveAs2 outputPath
        End If
    Next i
    
    ' Clean up
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit

Cleanup:
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Sub AutomateFinancialReporting()
    Dim wsSources As Worksheet
    Dim wsData As Worksheet
    Dim wsMetrics As Worksheet
    
    ' Define sheets for data entry and consolidation
    On Error GoTo ErrorHandler
    Set wsSources = ThisWorkbook.Sheets("DataSources")
    Set wsData = ThisWorkbook.Sheets("RawData")
    Set wsMetrics = ThisWorkbook.Sheets("CalculatedMetrics")
    
    ' Clear previous data from RawData sheet
    wsData.Cells.ClearContents

    ' Data consolidation from listed file paths
    Dim folderPath As String, filename As String
    Dim lastRow As Long
    Dim currentRow As Long
    currentRow = 2 ' Starting to enter data from row 2

    folderPath = wsSources.Range("A2").Value
    Do While folderPath <> ""
        ' Check if file exists
        If Dir(folderPath) = "" Then
            MsgBox "File not found: " & folderPath, vbExclamation
            folderPath = wsSources.Cells(currentRow, 1).Value
            Continue Do
        End If
        
        ' Open the workbook
        Dim wb As Workbook
        Set wb = Workbooks.Open(folderPath)
        
        ' Copy data from Sheet1 to RawData
        Dim sourceSheet As Worksheet
        Set sourceSheet = wb.Sheets(1)
        lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row

        ' Check if there is data to copy
        If lastRow < 2 Then
            MsgBox "No data found in: " & folderPath, vbExclamation
            wb.Close False
            folderPath = wsSources.Cells(currentRow, 1).Value
            Continue Do
        End If

        sourceSheet.Range("A2:F" & lastRow).Copy wsData.Cells(currentRow, 1)

        ' Increment currentRow for next data entry
        currentRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1
        
        ' Close the workbook
        wb.Close False
        
        ' Move to the next file path
        folderPath = wsSources.Cells(currentRow - 1, 1).Value
    Loop

    ' Metrics Calculation
    CalculateMetrics wsData, wsMetrics

    ' Formatting the summary report
    FormatSummaryReport wsMetrics

    ' Create summary charts
    CreateCharts wsMetrics
    
    ' Finish up
    MsgBox "Monthly Financial Report automation complete!"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Sub CalculateMetrics(wsData As Worksheet, wsMetrics As Worksheet)
    ' Example function to calculate financial metrics from RawData and output to CalculatedMetrics sheet
    Dim lastRow As Long
    On Error GoTo ErrorHandler
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Add headers for calculated metrics
    wsMetrics.Range("A1").Value = "Net Profit"
    wsMetrics.Range("B1").Value = "Revenue Growth (%)"
    wsMetrics.Range("C1").Value = "Profit Margin (%)"
    wsMetrics.Range("D1").Value = "Total Revenue"
    
    ' Loop to calculate net profit, revenue growth, profit margin, and total revenue
    Dim i As Long
    For i = 2 To lastRow
        ' Net Profit
        wsMetrics.Cells(i, 1).Value = wsData.Cells(i, 3).Value - wsData.Cells(i, 4).Value ' Revenue - Expenses
        
        ' Revenue Growth (%) - Assuming data is ordered by month
        If i > 2 Then
            wsMetrics.Cells(i, 2).Value = (wsData.Cells(i, 3).Value - wsData.Cells(i - 1, 3).Value) / wsData.Cells(i - 1, 3).Value * 100
        Else
            wsMetrics.Cells(i, 2).Value = "N/A" ' No previous data to calculate growth
        End If
        
        ' Profit Margin (%)
        If wsData.Cells(i, 3).Value <> 0 Then
            wsMetrics.Cells(i, 3).Value = (wsMetrics.Cells(i, 1).Value / wsData.Cells(i, 3).Value) * 100
        Else
            wsMetrics.Cells(i, 3).Value = "N/A"
        End If
        
        ' Total Revenue
        wsMetrics.Cells(i, 4).Value = wsData.Cells(i, 3).Value
    Next i
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Sub FormatSummaryReport(wsMetrics As Worksheet)
    ' Format the summary report in the CalculatedMetrics sheet
    On Error GoTo ErrorHandler
    With wsMetrics
        .Range("A1:D1").Font.Bold = True
        .Columns("A:D").AutoFit
        ' Add table formatting
        .ListObjects.Add(xlSrcRange, .Range("A1:D" & wsMetrics.Cells(wsMetrics.Rows.Count, "A").End(xlUp).Row), , xlYes).TableStyle = "TableStyleMedium9"
    End With
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Sub CreateCharts(wsMetrics As Worksheet)
    ' Create a sample chart from the metrics
    On Error GoTo ErrorHandler
    Dim chart As ChartObject
    Set chart = wsMetrics.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
    With chart.chart
        .SetSourceData Source:=wsMetrics.Range("A1:B" & wsMetrics.Cells(wsMetrics.Rows.Count, "A").End(xlUp).Row) ' Update range to cover all rows with data
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Monthly Revenue Growth"
    End With
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub


