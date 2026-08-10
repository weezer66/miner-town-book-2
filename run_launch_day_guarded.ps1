$ErrorActionPreference = 'Stop'
$kit = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $kit

$bat = Get-Content '.\Execute Launch Day.bat'
$smtpUserLine = ($bat | Select-String -Pattern 'set "SMTP_USER=' | Select-Object -First 1).Line
$smtpRawPassLine = ($bat | Select-String -Pattern 'set "SMTP_RAW_PASS=' | Select-Object -First 1).Line
$pyExeLine = ($bat | Select-String -Pattern 'set "PYEXE=' | Select-Object -First 1).Line

if (-not $smtpUserLine -or -not $smtpRawPassLine) {
    throw 'Could not find SMTP credentials in Execute Launch Day.bat'
}

$env:SMTP_USER = ($smtpUserLine -replace '^.*SMTP_USER=', '') -replace '"$',''
$rawPass = (($smtpRawPassLine -replace '^.*SMTP_RAW_PASS=', '') -replace '"$','')
$env:SMTP_PASS = $rawPass -replace ' ', ''

$pyRaw = if ($pyExeLine) {
    (($pyExeLine -replace '^.*PYEXE=', '') -replace '"$','')
} else {
    '%LOCALAPPDATA%\Programs\Python\Python313\python.exe'
}
$py = [Environment]::ExpandEnvironmentVariables($pyRaw)

if (-not (Test-Path $py)) {
    throw "Python not found at $py"
}

$campaignId = "launch-{0}" -f (Get-Date -Format 'yyyy-MM-dd')
$globalSkipDays = 30

Write-Output "Starting guarded launch-day send for campaign: $campaignId"
& $py 'launch_day_campaign.py' '--send' '--campaign-id' $campaignId '--include-responded' 'no' '--standalone' '--cover-image' 'C:\Users\Prisha\OneDrive\Documents\Miner Town\Archive\Miner_Town_Cover Page.png' '--cover-gif' 'C:\Users\Prisha\OneDrive\Documents\Miner Town\Archive\miner_town_cover_loop.gif' '--subject' 'My new novel, Miner Town: Awakening, is out now' '--limit' '50' '--daily-limit' '100' '--sleep-min' '20' '--sleep-max' '45' '--global-skip-days' $globalSkipDays
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Output 'LAUNCH_DAY_GUARDED_SEND_COMPLETE'
