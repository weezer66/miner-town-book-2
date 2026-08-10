# Paced sending — `run_burst.ps1`

Sends follow-ups in small batches with a cooldown between each, and **stops the
moment anything looks wrong** instead of quietly spinning through wasted rounds.

## Why this exists

On 2026-07-15 a network outage killed two rounds in the middle of a 40-send run.
The loop kept going because a connection failure surfaced as `0 sent, 0 errors`
and slipped past the error guard — the same silent-failure pattern documented in
the QA case study. `run_burst.ps1` replaces the old inline loop and closes that gap.

## Guardrails (all verified)

| Guard | Trigger | Result |
|-------|---------|--------|
| 1. Network preflight | `imap.gmail.com:993` unreachable (3 tries, 15s apart) | Abort before sending the round |
| 2. Exit-code check | `execute_followups.py` exits non-zero (IMAP **or** SMTP connect failure) | Abort the run |
| 3a. Error burst | A round reports 3+ errors | Abort (likely throttle) |
| 3b. Zero-send stop | A SEND round delivers nothing (queue empty / all held / all blocked) | Stop — nothing left to send |

On any abort, unsent rows keep `Follow-up? = Yes` and stay queued for next time.
Nothing is lost, nothing is double-sent.

## Usage

Credentials come from the environment and are **never** stored in the script:

```powershell
$env:SMTP_USER = 'saxena.ankit123@gmail.com'
$env:SMTP_PASS = '<gmail app password>'

# rehearse the whole harness without sending a single email:
.\run_burst.ps1 -Total 8 -Batch 4 -CooldownMinutes 0 -DryRun

# real run: 40 emails, 5 per round, 20 minutes between rounds:
.\run_burst.ps1 -Total 40 -Batch 5 -CooldownMinutes 20
```

Parameters: `-Total` (required), `-Batch` (default 5), `-CooldownMinutes`
(default 20), `-MinGapDays` (default 7), `-PreflightRetries` (default 3),
`-DryRun` (switch).

Each round still runs `execute_followups.py`, which does its own pre- and
post-send refresh, all existing guardrails (never-follow-up, has-book, cooldown,
repeat-content blocking), and the 120s IMAP/SMTP socket timeouts.

Every run writes a timestamped log to `burst_logs/`.
