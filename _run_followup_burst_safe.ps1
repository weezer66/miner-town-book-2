$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$bat = Get-Content '.\Execute Follow-ups.bat'
$userLine = ($bat | Select-String -Pattern 'set "SMTP_USER=' | Select-Object -First 1).Line
$passLine = ($bat | Select-String -Pattern 'set "SMTP_RAW_PASS=' | Select-Object -First 1).Line
if (-not $userLine -or -not $passLine) {
    throw 'SMTP lines missing in Execute Follow-ups.bat'
}

$env:SMTP_USER = (($userLine -replace '^.*SMTP_USER=', '') -replace '"$','')
$env:SMTP_PASS = ((($passLine -replace '^.*SMTP_RAW_PASS=', '') -replace '"$','') -replace ' ','')

& '.\run_burst.ps1' -Total 56 -Batch 5 -CooldownMinutes 20 -MinGapDays 7
