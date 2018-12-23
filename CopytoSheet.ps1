$path = "C:\Users\P2824589\Documents\Source.xlsx" 
$Excel = New-Object -ComObject excel.application 
$Excel.visible = $false 
$Workbook = $excel.Workbooks.open($path) 
$Worksheet = $Workbook.WorkSheets.item(1) 
$worksheet.activate()  
$range = $WorkSheet.Range("A1:B1").EntireColumn #Taking all of col A and B
$range.Copy() | out-null 
$Worksheet = $Workbook.Worksheets.item(2) #set destination to second sheet
$Range = $Worksheet.Range("G1") #Note Capital R as destination
$Worksheet.Paste($range)  
$workbook.Save()  
$Excel.Quit() 
Remove-Variable -Name excel 
[gc]::collect() 
[gc]::WaitForPendingFinalizers()