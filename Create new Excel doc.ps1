
$Excel01 = New-Object -ComObject Excel.Application #Create our object for new Excel file

$Excel01.Visible = $True #Display file
$sheets = Read-Host -Prompt "Choose amount of sheets"
$Excel01.SheetsInNewWorkbook = $sheets #Set number of sheets

$Workbook01 = $Excel01.Workbooks.Add() #Create the workbook with specified sheets

$Worksheet01 = $Workbook01.Worksheets.Item($sheets-1+1) #Actually create sheet WHY DO I HAVE TO DO THIS +1-1 THING WTF

$Worksheet01.Activate() #Set Focus

$Workbook01.Save()
#$location = Get-Location $Workbook01
"Saved file"

