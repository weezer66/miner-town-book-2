$ErrorActionPreference = 'Stop'
$kit = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $kit

$bat = Get-Content '.\Execute Follow-ups.bat'
$smtpUserLine = ($bat | Select-String -Pattern 'set "SMTP_USER=' | Select-Object -First 1).Line
$smtpRawPassLine = ($bat | Select-String -Pattern 'set "SMTP_RAW_PASS=' | Select-Object -First 1).Line
$pyExeLine = ($bat | Select-String -Pattern 'set "PYEXE=' | Select-Object -First 1).Line

if (-not $smtpUserLine -or -not $smtpRawPassLine) {
    throw 'Could not find SMTP credentials in Execute Follow-ups.bat'
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

Write-Output 'Starting refresh (no sending): response_report.py -> build_control_sheet.py'
& $py response_report.py
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& $py build_control_sheet.py
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Output 'REFRESH_COMPLETE'
