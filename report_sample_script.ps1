$file = "path"
$sheetName = "Sheet1"
$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $true
$objExcel.DisplayAlerts = $false
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.WorkSheets.Item("Sheet1")


$rowMax = ($sheet.UsedRange.Rows).count
$rowLamount,$colLamount = 1,2
$rowIR,$colIR = 1,4
$rowEMI_amount,$colEMI_amount = 1,5
$rowPaid,$colPaid = 1,7
$rowTerm,$colTerm = 1,8
$rowPamount,$colPamount = 1,9
$rowIamount,$colIamount = 1,10
for ($i=1; $i -le $rowMax-1; $i++)
{
$EMI = $sheet.Cells.Item($rowEMI_amount+$i,$colEMI_amount).text
$EMI = $EMI -as [int]
$Paid = $sheet.Cells.Item($rowPaid+$i,$colPaid).text
$Paid = $Paid -as [int]
$Term = $Paid/$EMI -as [int]
$sheet.Cells.Item($rowTerm+$i,$colTerm)= $Term
}
for ($i=1; $i -le $rowMax-1; $i++)
{
$IR = $sheet.Cells.Item($rowIR+$i,$colIR).text
$IR = $Ir -as [double]
$EMI = $sheet.Cells.Item($rowEMI_amount+$i,$colEMI_amount).text
$EMI = $EMI -as [double]
$temp_amount = $sheet.Cells.Item($rowLamount+$i,$colLamount).text
$temp_amount = $temp_amount -as [double]
$term = $sheet.Cells.Item($rowTerm+$i,$colTerm).text
$term = $term -as [int]
$jump=0
for ($j=0; $j -lt $term; $j++)
{
$int_amount = $IR/1200 * $temp_amount
$jump=$jump + $j
$sheet.Cells.Item($rowPamount,$colPamount+$jump)= "Principal_amount_" + ($j+1)
$sheet.Cells.Item($rowPamount+$i,$colPamount+$jump)= $EMI - $int_amount
$sheet.Cells.Item($rowIamount,$colIamount+$jump)= "Interest_amount" + ($j+1)
$sheet.Cells.Item($rowIamount+$i,$colIamount+$jump)= $int_amount
$jump=$j+1
$temp_amount= $temp_amount - $int_amount
}
}
$ext=".xlsx"
$path="path$ext"
$workbook.SaveAs($path) 
$workbook.Close



$objExcel.Workbooks.Close()
$objExcel.Quit()
Get-Process excel | Stop-Process -Force
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
