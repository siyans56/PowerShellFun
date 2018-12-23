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


function getdestination { #Function for getting the existing excel file
    "Please choose where to paste data: "
    Start-Sleep -s .5
    #$path = "C:\Users\P2824589\Documents\Source.xlsx" #Static to use for test
    $dpath = Get-FileName -initialDirectory "c:fso" #Get source file from dialog box
    "File chosen: $dpath" #Print file
    $PWorkbook = $excel.Workbooks.open($dpath) #Open source
    $PWorksheet = $PWorkbook.WorkSheets.item(1) #Variable for worksheet to use (also looking at first sheet)
    $Pworksheet.activate()  #Select the active sheet with variable    
    return $PWorkbook
}

function copypasteoldfile { #Function to copy/paste to existing file
    previewcols
    "Column names listed above."     
    $search=Read-Host -Prompt "Please input column name (if exact name unknown surround partial match with * (exmaple: *bill*)) or /quit to close: "
    if ($WorkSheet.Cells.Find($search)) #if the search element is found 
    {
        "Column $search found."
        Start-Sleep -Seconds .5  
        $WorkSheet.Cells.Find($search).EntireColumn.Copy() | out-null #Copy, hide return val

        #NEED TO SEPARATE INTO COPY PASTE FUNC AND SEARCH FUNC
        $PasteColumn = Read-Host -Prompt "Please enter column index to paste to (A1, B1, C1, ...): "  #Prompt for range
        $sheet = Read-Host -Prompt "Which sheet do you want to paste into? First sheet is 1, 2, and so on. Sheet must already exist to paste."  
        $Worksheet = $Workbook.Worksheets.item($sheet-1+1) #set destination to second sheet
        $Range = $Worksheet.Range($PasteColumn) #Set range as defined by user
        $Worksheet.Paste($range)  #Paste command
        "Successfully pasted column with $search" #confirmation
        #$Continue = Read-Host -Prompt "Have another column to paste? Press enter to continue or /quit to close: "
        #while ($Continue -ne "/quit") {copypaste}
    }
        else {
            "Could not find column $search."
            #while ($search -ne "/quit") {copypaste} #If dont quit, rerun
             } #>
}


#Body --------------------------------------------------------------------------------------------------------------------------


$FuncType = Read-Host -Prompt "Hit ENTER to paste to a new sheet, or 1 and ENTER to paste to an existing sheet: "


    "Copying to existing file"
    copypastesheet #First time execution

    while ((Read-Host -Prompt "Have another column to paste? Press enter to continue or /quit to close: ")-ne "/quit") {    
        copypastesheet
        } #Continued execution


#Cleanup ----------------------------------------------------------------------------------------------------------------------
$workbook.Save() #Save file
"File saved, closing..." 
Start-Sleep -Seconds 1  #Delay
$Excel.Quit()  #Quit
"Goodbye"
Remove-Variable -Name excel 
[gc]::collect() 
[gc]::WaitForPendingFinalizers()

