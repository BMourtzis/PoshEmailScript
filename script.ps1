# Write-Host "Hello World!" -ForegroundColor Blue
# $test = Read-Host -Prompt "Enter something"
# Write-Host "You wrote" $test

Write-Host "Read Excel file"

$ExcelObj = New-Object -comobject Excel.Application
$ExcelObj.visible = $false

$excelFile = "list.xlsx"
$ExcelWorkBook = $ExcelObj.Workbooks.Open(-join($pwd, "\", $excelFile))
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item(1)

[int]$lastRowValue = ($ExcelWorkSheet.UsedRange.rows.count + 1) - 1
[int]$lastColumnValue = ($ExcelWorkSheet.UsedRange.columns.count + 1) - 1

Write-Host $lastRowValue
Write-Host $lastColumnValue

# Get file headers
# Write-Host $ExcelWorkSheet.Range("A1:B1").value2
Write-Host $ExcelWorkSheet.Rows(1).Columns.Item(1).value2

#Get first line
Write-Host $ExcelWorkSheet.Range("A2:B2").value2

$ExcelWorkBook.close($true)
$ExcelObj.Quit()

$key = "κωδικός"
$value = "3"

$dictionary = @{email=1;file=2;}
$dictionary.Add($key, $value)
$dictionary.Item("email")
$dictionary.Item($key)