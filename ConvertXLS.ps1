# XLS to XLSX Batch convert script
# Forked from https://gist.github.com/gabceb/954418
# Works well using Office 365

$folderpath = Read-Host "Please entire file path"
$filetype ="*xls"
$convertErr = $false
$Log = @()
$oldPath = @()

Add-Type -AssemblyName Microsoft.Office.Interop.Excel

$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$excel = New-Object -ComObject excel.application
$excel.visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false

Get-ChildItem -Path $folderpath -Include $filetype -recurse | Where-Object {$_.FullName -notlike '*\old\*'} |
ForEach-Object {
	$path = ($_.fullname).substring(0, ($_.FullName).lastindexOf("."))
	$oldPath = $path + ".xls"
	$convertErr = $false

	try
	{
		$excel.visible = $false
	}
	catch
	{
		
	}

	write-host "`r`nConverting $oldPath"
	$path += ".xlsx"

	try
	{
		$workbook = $excel.workbooks.open($_.fullname,0,0,5,$null)
		$workbook.saveas($path, $xlFixedFormat)
		$workbook.close()
		write-host "$path successfully converted"
	}
	catch
	{
		write-host "Error at $path `r`n"
		$Log += ("Error at $path")
		$convertErr = $true
	}
	
	$oldFolder = $path.substring(0, $path.lastIndexOf("\")) + "\old"
	
	if ($convertErr -ne $true)
	{
		#write-host $oldFolder
		if(-not (test-path $oldFolder))
		{
			new-item $oldFolder -type directory
		}

		$Log += $path + " successfully converted"

		move-item -LiteralPath $_.fullname $oldFolder
		write-host "$oldPath moved to $oldFolder `r`n"
		$Log += "$oldPath moved to $oldFolder `r`n"
	}
}

$Log | Out-File -FilePath ($folderpath + "\outputXLSX" + (Get-Date -Format "ddmmyyyyHHmm") + ".txt")

$excel.Quit()
$excel = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
