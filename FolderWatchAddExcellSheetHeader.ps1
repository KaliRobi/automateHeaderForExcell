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
            $ExcelInstance.DisplayAlerts = $false
            $ExcelInstance.visible=$false
            
            #The folder where this ps1 file is located. If doesnt work for some reason the path can be adjusted to the folder where SQL developer exports the .xlsx files between " "
            $Folder = "$($pwd)"
            
            #  the script selects the last created file in the directory with .xlsx extension
            $LatestFile = (Get-ChildItem -Path $Folder -Filter *.xlsx -File | Sort-Object -Descending -Property CreationTime | select -First 1 )
            $LatestFile.isReadOnly = $false
            
            # show the selected file in the terminal for clarity.
            echo "$($LatestFile.Name) was selected"
             
            $CurrentExcelWorkBook = $ExcelInstance.Workbooks.Open("$($Folder)\\$($LatestFile.Name)", $null, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true  )
            
            $WorksheetSheet =  $CurrentExcelWorkBook.Sheets.Item('Export Worksheet')
            
            $filterRange = $WorksheetSheet.UsedRange
            
            $columnCount = $WorksheetSheet.UsedRange.Columns.Count
            
            #sets columns' filters on
            $filterRange.AutoFilter(1)
            #header text will be bold
            $WorksheetSheet.Rows.Item(1).Font.Bold = $true
            
            # column width will be adjusted based on the text
            for ($i = 1; $i -le $columnCount; $i++) {
                $WorksheetSheet.Columns($i).AutoFit() | Out-Null
                #progress bar
                Write-Output "--------------"
            }
            
            echo "script run all is ok, closing process"
            #cleanup
            $CurrentExcelWorkBook.Close($true) 
            $ExcelInstance.Quit() 
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelInstance) | Out-Null
}


Register-ObjectEvent -InputObject $watcher -EventName Created -Action $action
