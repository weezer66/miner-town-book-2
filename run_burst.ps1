<#
.SYNOPSIS
  Paced follow-up sender: send in small batches with cooldowns between them,
  and STOP the moment anything looks wrong instead of quietly wasting rounds.

.DESCRIPTION
  A network outage once killed two rounds mid-run while the loop kept spinning,
  because a connection failure surfaced as "0 sent, 0 errors" and slipped past
  the error guard. This script closes that gap with three hard stops:
    1. Network preflight before every round (skip the whole run if Gmail is
       unreachable, rather than firing a round into a dead socket).
    2. Exit-code check: execute_followups.py exits non-zero on any connection
       failure, so a non-zero code aborts immediately.
    3. Zero-send stop: a SEND round that delivers nothing (queue empty, all
       held/blocked, or a silent connection issue) ends the run.

  Credentials are read from the environment (SMTP_USER / SMTP_PASS) and are
  NEVER stored in this file. Set them in the session before running.

.EXAMPLE
  $env:SMTP_USER='you@gmail.com'; $env:SMTP_PASS='app-password'
  .\run_burst.ps1 -Total 40 -Batch 5 -CooldownMinutes 20
  .\run_burst.ps1 -Total 8 -Batch 4 -CooldownMinutes 0 -DryRun   # rehearsal, no mail
#>
param(
  [Parameter(Mandatory = $true)][int]$Total,
  [int]$Batch = 5,
  [int]$CooldownMinutes = 20,
  [int]$MinGapDays = 7,
  [int]$PreflightRetries = 3,
  [switch]$DryRun
)

$kit      = $PSScriptRoot
$py       = "C:\Users\Prisha\AppData\Local\Programs\Python\Python313\python.exe"
$xlsx     = Join-Path $kit 'followups.xlsx'
$imapHost = 'imap.gmail.com'
$mode     = if ($DryRun) { 'DRY' } else { 'SEND' }

# --- logging ---------------------------------------------------------------
$logDir = Join-Path $kit 'burst_logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Force $logDir | Out-Null }
$log = Join-Path $logDir ("burst_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
function Say([string]$m) {
  $line = "[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $m
  Write-Host $line
  try { Add-Content -Path $log -Value $line -Encoding utf8 -ErrorAction Stop } catch {}
}
function Log([string]$block) {
  try { Add-Content -Path $log -Value $block -Encoding utf8 -ErrorAction Stop } catch {}
}

# --- credentials must come from the environment, never this file -----------
if (-not $env:SMTP_USER -or -not $env:SMTP_PASS) {
  Say "ABORT: SMTP_USER / SMTP_PASS not set in the environment."
  exit 2
}

# --- network preflight (the guard the outage exposed) ----------------------
function Test-GmailReachable {
  for ($i = 1; $i -le $PreflightRetries; $i++) {
    try {
      $client = [System.Net.Sockets.TcpClient]::new()
      $async = $client.BeginConnect($imapHost, 993, $null, $null)
      $ok = $async.AsyncWaitHandle.WaitOne(8000)
      if ($ok -and $client.Connected) {
        $client.EndConnect($async) | Out-Null
        $client.Close()
        return $true
      }
      $client.Close()
    } catch { }
    if ($i -lt $PreflightRetries) { Start-Sleep -Seconds 15 }
  }
  return $false
}

Set-Location $kit
$maxRounds = [math]::Ceiling($Total / [double]$Batch)
Say "=== BURST START : target=$Total batch=$Batch cooldown=${CooldownMinutes}m mode=$mode (max $maxRounds rounds) ==="

$totalSent = 0
for ($r = 1; $r -le $maxRounds; $r++) {

  $limit = [math]::Min($Batch, $Total - $totalSent)
  if ($limit -le 0) { Say "target reached"; break }

  # sheet must be closed or the rebuild inside execute will fail
  try { $fs = [System.IO.File]::Open($xlsx, 'Open', 'ReadWrite', 'None'); $fs.Close() }
  catch { Say "ABORT round ${r}: followups.xlsx is LOCKED (close Excel)."; break }

  # GUARD 1: do not fire a round into a dead network
  if (-not (Test-GmailReachable)) {
    Say "ABORT round ${r}: ${imapHost}:993 unreachable after $PreflightRetries tries. Remaining rows stay queued."
    break
  }

  Say "--- ROUND $r/$maxRounds : $mode $limit ---"
  $argv = @('execute_followups.py', '--limit', "$limit", '--min-gap-days', "$MinGapDays")
  if (-not $DryRun) { $argv = @('execute_followups.py', '--send', '--limit', "$limit", '--min-gap-days', "$MinGapDays") }

  $out  = & $py @argv 2>&1 | Out-String
  $code = $LASTEXITCODE
  Log $out

  # GUARD 2: python exits non-zero on a fatal send-connection failure (IMAP or
  # SMTP). We trust the exit code plus the fatal-only marker 'Could not connect
  # to Gmail' — do NOT match generic socket phrases like 'forcibly closed' or
  # 'getaddrinfo failed'. Those also appear in a NON-fatal refresh warning
  # (e.g. WinError 10054 during a pre/post-send Gmail rescan), which
  # execute_followups.py catches and continues past — sending still succeeds
  # and it exits 0. Matching them here caused a spurious abort/undercount.
  if ($code -ne 0 -or ($out -match 'Could not connect to Gmail')) {
    Say "ABORT round ${r}: connection/hard failure (exit=$code). Remaining rows stay queued."
    break
  }

  $sent = 0; $errors = 0
  if     ($out -match 'Sent:\s*(\d+)')       { $sent   = [int]$Matches[1] }
  elseif ($out -match 'Would send:\s*(\d+)') { $sent   = [int]$Matches[1] }
  if     ($out -match 'Errors:\s*(\d+)')     { $errors = [int]$Matches[1] }
  $totalSent += $sent
  Say "ROUND $r result -> sent=$sent errors=$errors (running total=$totalSent)"

  # GUARD 3a: a burst of errors looks like a throttle
  if ($errors -ge 3) { Say "ABORT: $errors errors this round - likely throttle. Remaining rows stay queued."; break }

  # GUARD 3b: nothing sent means nothing left to send this run - do not spin
  if ($sent -le 0) { Say "STOP round ${r}: nothing sendable (queue empty / all held / all blocked). Remaining rows stay queued."; break }

  if ($totalSent -ge $Total) { Say "target $Total reached"; break }
  if ($r -lt $maxRounds) { Say "cooldown $CooldownMinutes min ..."; Start-Sleep -Seconds ($CooldownMinutes * 60) }
}

Say "=== BURST END : total $mode = $totalSent of $Total ==="
