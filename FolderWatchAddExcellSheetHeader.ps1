#the location where SQL developer exports
$folderToWatch = ''


$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $folderToWatch
$watcher.Filter = "*.*"
$watcher.EnableRaisingEvents = $true
$watcher.IncludeSubdirectories = $false

#action to do
$action = {
            
        $ExcelInstance = New-Object -ComObject excel.application

        $ExcelInstance.visible=$false

        # adjust to your needs, the ps1 files needs to be where the SQL developer exports the .xlsx file
        $Folder = "$($pwd)"

        #  the script selects the last created file in the directory
        $LatestFile = (Get-ChildItem -Attributes !Directory -Path $Folder | Sort-Object -Descending -Property CreationTime | select -First 1)
        echo "$($LatestFile.Name) was selected"
        $CurrentExcelWorkBook = $ExcelInstance.Workbooks.Open("$($Folder)\\$($LatestFile.Name)"  )

        $WorksheetSheet =  $CurrentExcelWorkBook.Sheets.Item('Export Worksheet')

        $filterRange = $WorksheetSheet.UsedRange
        $filterRange.AutoFilter(1)

        $columnCount = $WorksheetSheet.UsedRange.Columns.Count
        $WorksheetSheet.Rows.Item(1).Font.Bold = $true

        for ($i = 1; $i -le $columnCount; $i++) {
            $WorksheetSheet.Columns($i).AutoFit()
        }

        echo "script run all is ok, closing process"
        $CurrentExcelWorkBook.Close($true) 
        $ExcelInstance.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelInstance) | Out-Null
}


Register-ObjectEvent -InputObject $watcher -EventName Created -Action $action
