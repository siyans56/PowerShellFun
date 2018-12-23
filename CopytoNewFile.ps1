"Script by Siyan Shaikh, adopted from various sources and research online"
<#
TO DO:
Let user pick file paths using popup box DONE
Let user copy column from multiple files in one folder
Have console print column names so user can choose without opening file
Let user type in which columns they want copied (enter each name until done) DONE
Let user choose if they want to paste to new sheet or new workbook or existing sheet
If pasting to new book, let user create how many sheets and also pick which sheet
#>

#Initialization --------------------------------------------------------------------------------------------------------------------

"Please choose where the source data is coming from: "
Start-Sleep -s .5
#$path = "C:\Users\P2824589\Documents\Source.xlsx" #Static to use for test
$spath = Get-FileName -initialDirectory "c:fso" #Get source file from dialog box
"File chosen: $spath" #Print file

$Excel = New-Object -ComObject excel.application 
$Excel.visible = $True #make it so the file doesn't visibly open everytime
$Workbook = $excel.Workbooks.open($spath) #Open source
$Worksheet = $Workbook.WorkSheets.item(1) #Variable for worksheet to use (also looking at first sheet)
$worksheet.activate()  #Select the active sheet with variable

#Function Definitions -------------------------------------------------------------------------------------------------------------

Function Get-FileName($initialDirectory) #Function made by user ScriptingGuy1 on Microsoft TechNet
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

function initnewfile { #Function for creating the new file

    #$Pastebook = New-Object -ComObject Excel.Application #Create our object for new Excel file
    #$Pastebook.Visible = $True #Display file
    $numsheet = Read-Host -Prompt "How many sheets do you want in the new file?: "
    $Excel.SheetsInNewWorkbook = $numsheet #Set number of sheets for all files made in this object
    $Newbook = $Excel.Workbooks.Add() #Create the workbook with specified sheets
    $Newsheet = $Newbook.Worksheets.Item($numsheet-1+1) #Actually create sheet
    $Newsheet.Activate() #Set Focus
    return $Newbook | Out-Null
}

Function previewcols {
  <# for ($i=1; $i -le 50; $i++) {
        Write-Host -NoNewLine $WorkSheet.Cells.Item(1,$i).Text ", "
        $length = $WorkSheet.Cells.Item(1,$i).Text.length
        if ($length -eq 0) {break}
    } THIS IS IF YOU WNAT IT ALL ON ONE LINE#>
    for ($i=1; $i -le 50; $i++) {
    Write-Host $WorkSheet.Cells.Item(1,$i).Text
    $length = $WorkSheet.Cells.Item(1,$i).Text.length
    if ($length -eq 0) {break}
    }
}

function copypastenewfile($Newbook) { #Function for copying and pasting to a new excel file
    previewcols
    $search=Read-Host -Prompt "Please input column name (if exact name unknown surround partial match with * (exmaple: *bill*)) or /quit to close: "
    if ($WorkSheet.Cells.Find($search)) #if the search element is found 
    {
        "Column $search found."
        Start-Sleep -Seconds .5  
        $WorkSheet.Cells.Find($search).EntireColumn.Copy() | out-null #Copy, hide return val
        $sheet = Read-Host -Prompt "Which sheet do you want to paste into? First sheet is 1, second is 2, and so on. Sheet must already exist to paste."
        $PasteColumn = Read-Host -Prompt "Please enter column index to paste to (A1, B1, C1, ...): "  #Prompt for range
        $Dsheet = $Newbook.Worksheets.Item($sheet-1+1) #Select sheet to paste into
        $Range = $Dsheet.Range($PasteColumn) #Set range as defined by user
        $Dsheet.Paste($Range)  #Paste command
        "Successfully pasted column with $search" #confirmation
        $Newbook.Save()
    }
        else {
            "Could not find column $search."
            }
}

function altfunction { #Pls work
    previewcols
    "Column names listed above." 
    $search=Read-Host -Prompt "Please input column name (if exact name unknown surround partial match with * (exmaple: *bill*)) or /quit to close: "
    if ($WorkSheet.Cells.Find($search)) #if the search element is found 
    {
        "Column $search found."
        Start-Sleep -Seconds .5  
        $WorkSheet.Cells.Find($search).EntireColumn.Copy() | out-null #Copy, hide return val
        initnewfile
        $sheet = Read-Host -Prompt "Which sheet do you want to paste into? First sheet is 1, second is 2, and so on. Sheet must already exist to paste."
        $PasteColumn = Read-Host -Prompt "Please enter column index to paste to (A1, B1, C1, ...): "  #Prompt for range
        $Dsheet = $Newbook.Worksheets.Item($sheet-1+1) #Select sheet to paste into
        $Range = $Dsheet.Range($PasteColumn) #Set range as defined by user
        $Dsheet.Paste($Range)  #Paste command
        "Successfully pasted column with $search" #confirmation
        $Newbook.Save()
    }
        else {
            "Could not find column $search."
            }    
}

#Body --------------------------------------------------------------------------------------------------------------------------



"Copying to new file"
$Newbook = initnewfile
copypastenewfile $Newbook
while ((Read-Host -Prompt "Have another column to paste? Press enter to continue or /quit to close: ")-ne "/quit") {    
    copypastenewfile $Newbook
    } #Continued execution
    
#"Running"
#altfunction

#Cleanup ----------------------------------------------------------------------------------------------------------------------
$workbook.Save() #Save file
"File saved, closing..." 
Start-Sleep -Seconds 1  #Delay
$Excel.Quit()  #Quit
"Goodbye"
Remove-Variable -Name excel 
[gc]::collect() 
[gc]::WaitForPendingFinalizers()

