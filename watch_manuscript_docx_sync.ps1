$root = "C:/Users/Prisha/OneDrive/Documents/Miner Town/Book 2"
$manuscriptDir = "$root/manuscript"
$syncScript = "$root/sync_manuscript_docx.py"
$pythonExe = "C:/Users/Prisha/AppData/Local/Programs/Python/Python313/python.exe"

function Invoke-Sync {
    Set-Location $root
    & $pythonExe $syncScript --dir $manuscriptDir
}

Invoke-Sync

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $manuscriptDir
$watcher.Filter = "*.md"
$watcher.EnableRaisingEvents = $true

Register-ObjectEvent $watcher "Changed" -Action {
    Start-Sleep -Milliseconds 250
    Set-Location $root
    & $pythonExe $syncScript --dir $manuscriptDir
    Write-Host "Synced manuscript DOCX files at $(Get-Date -Format o)"
} | Out-Null

Register-ObjectEvent $watcher "Created" -Action {
    Start-Sleep -Milliseconds 250
    Set-Location $root
    & $pythonExe $syncScript --dir $manuscriptDir
    Write-Host "Synced manuscript DOCX files at $(Get-Date -Format o)"
} | Out-Null

Register-ObjectEvent $watcher "Renamed" -Action {
    Start-Sleep -Milliseconds 250
    Set-Location $root
    & $pythonExe $syncScript --dir $manuscriptDir
    Write-Host "Synced manuscript DOCX files at $(Get-Date -Format o)"
} | Out-Null

Write-Host "Watching manuscript markdown for DOCX sync..."
Write-Host "Press Ctrl+C to stop."
while ($true) { Start-Sleep -Seconds 1 }
