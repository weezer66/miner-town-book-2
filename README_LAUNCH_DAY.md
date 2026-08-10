# Launch Day Email Mechanism

This flow sends launch-day emails to contacts already in your database in controlled daily chunks, with no duplicates.

## What it uses
- Recipient source: `followups.xlsx` -> `Follow-ups` sheet.
- Default template: `templates/launch-reminder.txt` and `templates/launch-reminder.html`.
- Exclusions: `suppression.txt` and (by default) rows where `Never Follow-up? = yes`.
- Progress log: `launch_day_sent_log.txt`.
- Default audience: non-responders first (`--include-responded no`).
- Default mode: standalone launch subject (not threaded replies).

## Safety built in
- Dry run by default.
- Requires `--send` to actually deliver.
- Per-run cap with `--limit` (default 50).
- Per-day cap with `--daily-limit` (default 100).
- Random pacing delay (default 20-45 seconds between sends).
- Resume across days: already-sent recipients for the same campaign are skipped automatically.

## Quick start (Windows)
1. Double-click `Execute Launch Day.bat`.
2. Review preview output.
3. Type `SEND` to deliver only that batch.
4. Run again later the same day or next day until queue is finished.

## Phase 2 preset (responders only, gated)
- Use `Execute Launch Day Phase 2.bat` for responders.
- This preset is locked until non-responders are exhausted for the same campaign ID.
- If non-responders are still pending, Phase 2 prints a lock message and sends nothing.

## Command-line usage
Preview:

```powershell
$env:SMTP_USER = "your@gmail.com"
$env:SMTP_PASS = "your-16-char-app-password"
python launch_day_campaign.py --campaign-id launch-2026-07-28 --include-responded no --standalone --subject "MINER TOWN: AWAKENING is live today" --limit 50 --daily-limit 100 --sleep-min 20 --sleep-max 45
```

Send:

```powershell
python launch_day_campaign.py --campaign-id launch-2026-07-28 --include-responded no --standalone --subject "MINER TOWN: AWAKENING is live today" --limit 50 --daily-limit 100 --sleep-min 20 --sleep-max 45
python launch_day_campaign.py --send --campaign-id launch-2026-07-28 --include-responded no --standalone --subject "MINER TOWN: AWAKENING is live today" --limit 50 --daily-limit 100 --sleep-min 20 --sleep-max 45
```

## Useful options
- `--include-responded all|yes|no`
  - `all` = everyone in `followups.xlsx`
  - `yes` = only those who have replied before
  - `no` = only no-response contacts
- `--phase-2`
  - Requires `--include-responded yes`.
  - Enforces: responders can only start when non-responders are exhausted for that campaign ID.
- `--include-never`
  - Includes rows marked `Never Follow-up?=yes` (off by default)
- `--threaded`
  - Override and send threaded replies (default mode is standalone launch subject).

## Placeholder fields in template
The template supports:
- `{first_name}`
- `{booksprout_link}`
- `{launch_date}`
- `{launch_day}`

If you do not use a placeholder, nothing breaks.

## Autonomous Day 2 agents
You now have three custom agents in `.github/agents` for hands-off Day 2 operations:

- `miner-town-day2-safety-gate`
  - Runs strict readiness checks (duplicates, lock/cooldown, active sender process, Phase 2 gating).
  - Returns GREEN/YELLOW/RED with exact recovery actions.

- `miner-town-day2-executor`
  - Runs the Day 2 workflow end-to-end: preflight, refresh, preview, optional guarded send, post-run verification.
  - Requires explicit `SEND DAY2` confirmation before any live send.

- `miner-town-weekly-opportunity-queue`
  - Builds a 7-day free opportunity submission queue with direct links, rule notes, and prewritten copy blocks.
  - Uses `channel_performance_tracker.csv` to prioritize what is most efficient.

Tracker file:
- `channel_performance_tracker.csv`
  - Scores each channel by effort-to-result and drives automatic prioritization.

Autonomy + KPI config:
- `campaign_autonomy_config.yaml`
  - Canonical campaign targets and constraints (30-day goals, user daily time cap, full-autonomy with guardrails).
- `daily_kpi_tracker.csv`
  - Daily rollup for sales and profile visits (Goodreads + Amazon Author Central) against 30-day targets.

Suggested usage pattern:
1. Ask Safety Gate to validate readiness for your lane (launch, phase 2, follow-up, combined).
2. Ask Weekly Opportunity Queue to generate this week's free submissions plan.
3. If GREEN, ask Day 2 Executor to run preview (it should consume queue + tracker priorities).
4. If preview looks correct, reply with `SEND DAY2` to authorize live execution.
5. At end of day, log outcomes in `daily_kpi_tracker.csv` so the next run auto-rebalances to stay on KPI pace.
