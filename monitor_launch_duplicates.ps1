param(
  [string]$CampaignId = "launch-2026-07-29",
  [string]$LogPath = ".\\launch_day_sent_log.txt"
)

$ErrorActionPreference = "Stop"

function Check-Duplicates {
  param(
    [string]$Cid,
    [string]$Path
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    Write-Output "[WARN] Log file not found: $Path"
    return
  }

  $rows = Import-Csv -LiteralPath $Path | Where-Object { $_.campaign_id -eq $Cid }
  $dupGroups = @($rows | Group-Object email | Where-Object { $_.Count -gt 1 } | Sort-Object Count -Descending)
  $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

  if ($dupGroups.Count -gt 0) {
    Write-Output "[ALERT] $stamp campaign=$Cid duplicates=$($dupGroups.Count)"
    $dupGroups | ForEach-Object {
      Write-Output ("[ALERT] email={0} count={1}" -f $_.Name, $_.Count)
    }
  }
  else {
    Write-Output "[OK] $stamp campaign=$Cid total=$($rows.Count) duplicates=0"
  }
}

$resolved = Resolve-Path -LiteralPath $LogPath
$watchDir = Split-Path -Parent $resolved
$watchFile = Split-Path -Leaf $resolved

Write-Output "Monitoring $watchFile for campaign $CampaignId"
Check-Duplicates -Cid $CampaignId -Path $resolved

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watchDir
$watcher.Filter = $watchFile
$watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size -bor [System.IO.NotifyFilters]::FileName
$watcher.EnableRaisingEvents = $true

Register-ObjectEvent -InputObject $watcher -EventName Changed -SourceIdentifier "launch_log_changed" | Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier "launch_log_created" | Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Renamed -SourceIdentifier "launch_log_renamed" | Out-Null

try {
  while ($true) {
    $evt = Wait-Event
    if (-not $evt) {
      continue
    }

    if ($evt.SourceIdentifier -in @("launch_log_changed", "launch_log_created", "launch_log_renamed")) {
      Remove-Event -EventIdentifier $evt.EventIdentifier -ErrorAction SilentlyContinue
      Check-Duplicates -Cid $CampaignId -Path $resolved
    }
    else {
      Remove-Event -EventIdentifier $evt.EventIdentifier -ErrorAction SilentlyContinue
    }
  }
}
finally {
  Unregister-Event -SourceIdentifier "launch_log_changed" -ErrorAction SilentlyContinue
  Unregister-Event -SourceIdentifier "launch_log_created" -ErrorAction SilentlyContinue
  Unregister-Event -SourceIdentifier "launch_log_renamed" -ErrorAction SilentlyContinue
  $watcher.Dispose()
}
