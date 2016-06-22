$excel_file_path = 'Q:\New Laptop Tracking\computerNameBook.xlsm' 
# Instantiate the COM object
$Excel = New-Object -ComObject Excel.Application
$Excel.DisplayAlerts = $false
$ExcelWorkBook = $Excel.Workbooks.Open($excel_file_path)
# First Sheet in Workbook
$ExcelWorkSheet = $Excel.Sheets.item(1)
$ExcelWorkSheet.activate()
# Gets username and computer name on the domain
$Name = [Environment]::UserName 
$ComputerName = $env:Computername
# Find the first empty row
$row = $ExcelWorkSheet.UsedRange.Rows.Count + 1
$ExcelWorkSheet.Cells.Item($row,1) = $Name
$ExcelWorkSheet.Cells.Item($row,2) = $ComputerName 
 
$ExcelWorkBook.Save()
$ExcelWorkBook.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
Stop-Process -Name EXCEL -Force