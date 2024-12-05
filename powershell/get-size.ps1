$allFiles = Get-ChildItem -Path '' -Filter '*.xlsx' -File
$currentDirectory = Get-Location

$Excel = New-Object -ComObject "Excel.Application" 
$Excel.Visible = $false #Runs Excel in the background. 
$Excel.DisplayAlerts = $false #Supress alert messages. 

foreach ($file in $allFiles) {
    # general size of file
    Write-Host "File: $file"
    $fileSizeBytes = $file.Length
    $fileSizeMB = $fileSizeBytes / 1MB  # Divide by 1MB to convert bytes to megabytes
    Write-Host "File size in MB: $fileSizeMB"
    # workbook split
    $Workbook = $Excel.Workbooks.open($file.FullName)
    $total = 0
    foreach($Worksheet in $Workbook.Worksheets) 
    { 
        $range = $Worksheet.UsedRange.Count
        $total = $total + $range
    } 
    foreach($Worksheet in $Workbook.Worksheets) 
    { 
        $range = $Worksheet.UsedRange.Count
        $sheetFileName = $Worksheet.Name
        Write-Output "Worksheet: $sheetFileName"
        $worksheetSize = $range * $fileSizeBytes / $total
        $worksheetSizeMB = ($worksheetSize / 1MB) -as [decimal]
        Write-Host "Worksheet size in MB: $worksheetSizeMB"
    } 
    $Workbook.Close()
}
$Excel.Quit() 