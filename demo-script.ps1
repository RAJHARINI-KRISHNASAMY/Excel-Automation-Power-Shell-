$objExcel = New-Object -ComObject Excel.Application
$sheetName = "Sheet1"
$objExcel.Visible = $true
$objExcel.DisplayAlerts = $false
$file = "path"
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$sheet2=$workbook.Worksheets.Item(2)
$rowMax = ($sheet.UsedRange.Rows).count
$rowLamount,$colLamount = 1,2
$total=0
for ($i=2; $i -le $rowMax-1; $i++)
{
$loan_amount = $sheet.Cells.Item($i,$colLamount).text
$loan_amount = $loan_amount -as [int]
$total=$loan_amount+$total
}
$sheet2.Cells.Item(1,1)= "Total Loan Amount"
$sheet2.Cells.Item(1,2)= $total
$path="path"
$workbook.SaveAs($path)
$workbook.Close
$objExcel.Workbooks.Close()
$objExcel.Quit()
Get-Process excel | Stop-Process -Force
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
