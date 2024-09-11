# DOC to DOCX Batch convert script
# Forked from https://gist.github.com/gabceb/954418
# Works well using Office 365

$folderpath = "K:\021-PDX-02"
$filetype ="*doc"
$convertErr = $false
$Log = @()
$oldPath = @()

Add-Type -AssemblyName Microsoft.Office.Interop.Word

$Format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument
$word = New-Object -ComObject Word.Application
$word.visible = $false

Get-ChildItem -Path $folderpath -Include $filetype -recurse | Where-Object {$_.FullName -notlike '*\old\*'} | 
ForEach-Object -Parallel {
	#Replace braces with parentheses
	if($_.Name.Contains("["))
	{
		Rename-Item -NewName { $_.Name -replace "[","(" }
	}
	if($_.Name.Contains("]"))
	{
		Rename-Item -NewName { $_.Name -replace "]",")" }
	}

	$path = ($_.fullname).substring(0, ($_.FullName).lastindexOf("."))
	$oldPath = $path + ".doc"
	$convertErr = $false

	write-host "`r`nConverting $oldPath"
	$path += ".docx"

	try
	{
		$document = $word.Documents.Open($_.FullName)
		$document.SaveAs($path,$Format)
		$document.Close()
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

		move-item $_.fullname $oldFolder
		write-host "$oldPath moved to $oldFolder `r`n"
		$Log += "$oldPath moved to $oldFolder `r`n"
	}
}

$Log | Out-File -FilePath ($folderpath + "\outputDOCX" + (Get-Date -Format "ddmmyyyyHHmm") + ".txt")

$word.Quit()
$word = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()