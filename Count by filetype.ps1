$path = Read-Host "Please entire file path"
#Get-Childitem $path -Recurse -Include "*.doc" | where { -not $_.PSIsContainer } | Where-Object {$_.FullName -notlike '*\old\*'} | group Extension -NoElement | sort count -desc
Get-Childitem $path -Recurse | where { -not $_.PSIsContainer } | Where-Object {$_.FullName -notlike '*\old\*'} | group Extension -NoElement | sort count -desc