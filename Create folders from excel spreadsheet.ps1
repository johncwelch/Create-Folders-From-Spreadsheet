#add frameworks to make a file selection browser
Add-Type -AssemblyName System.Windows.Forms

#create the file browser object, only allow tsv files to be seen, only sign a single file (default for mulstiselect is false)
#default starting directory is the user desktop
$fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
	InitialDirectory = [Environment]::GetFolderPath('Desktop') 
	RestoreDirectory = $true
	Filter = 'Excel (*.xlsx)|*.xlsx'
	Multiselect = $false
	ShowHelp = $false
} 

#display the file browser
$null = $fileBrowser.ShowDialog()

#set filepath to var
$sourceFile = $fileBrowser.FileName

#create new excel instance, open the spreadsheet and make it visible
$convertExcelToFolders = New-Object -ComObject Excel.Application
$convertExcelToFolders.Workbooks.Open($sourceFile)
#this step is optional, but handy for troubleshooting
#$convertExcelToFolders.Visible = $true

##Now we do the excel things
#set the active worksheet so we have something to act against
$convertExcelToFoldersSheet = $convertExcelToFolders.ActiveSheet

#get all the USED cells in the worksheet
$theNameListRange = $convertExcelToFoldersSheet.UsedRange

#get the number of cells in our range, we only care about the first column
$numCells = $theNameListRange.Columns.Item(1).rows.count
#create the ending cell. This is NOT the most elegant way to do it, but it's simple
#basically we concatentate the letter "A" with the number of cells. So we'll go from A1 to A(numcells)
$endCell = "A" + $numCells

##now we set up the destination. Allows for choosing a destination folder or creating a new one.
$FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
	RootFolder = $FolderBrowserDialog.RootFolder = 'Desktop'
	ShowNewFolderButton = $true
}

$null = $FolderBrowserDialog.ShowDialog()

#this is where the new folders will be created
$destinationFolder = $FolderBrowserDialog.SelectedPath

#now, let's go through our list of cells, grab the content of each cell, and make a directory with that as the name:

foreach($cell in $convertExcelToFoldersSheet.Range("A1:$endCell").Cells) {
	#not needed but makes things more clear. Value2 gets you the contents of the cell
	$foldername = $cell.Value2
	#create a new folder for every name in our list
	New-Item -path $destinationFolder -Name $foldername -ItemType "directory"
}


#now we quit the excel process we created. Unfortunately, just closing the file if visible doesn't do this.

$convertExcelToFolders.Workbooks.Close()
$convertExcelToFolders.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($convertExcelToFolders)