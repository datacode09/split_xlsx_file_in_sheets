# Define the path to your XLSX file
$filepath = "D:\powershell\test.xlsx"

# Create an Excel application object
$Excel = New-Object -ComObject "Excel.Application"
$Excel.Visible = $false # Run Excel in the background
$Excel.DisplayAlerts = $false # Suppress alert messages

# Open the workbook
$Workbook = $Excel.Workbooks.Open($filepath)
$WorkbookName = "test.xlsx"
$output_type = "xlsx"

if ($Workbook.Worksheets.Count -gt 0) {
    Write-Output "Now processing: $WorkbookName"
    $FileFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook
    $WorkbookName = $filepath -replace ".xlsx", ""

    foreach ($Worksheet in $Workbook.Worksheets) {
        $Worksheet.Copy()
        $ExtractedFileName = "$WorkbookName~~" + $Worksheet.Name + ".$output_type"
        $Excel.ActiveWorkbook.SaveAs($ExtractedFileName, $FileFormat)
        Write-Output "Created file: $ExtractedFileName"
        $Excel.ActiveWorkbook.Close
    }
}

# Clean up and close Excel
$Workbook.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
Stop-Process -Name EXCEL
Remove-Variable Excel
