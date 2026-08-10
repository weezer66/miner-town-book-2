# Miner Town Follow-up Kit — How It Flows

A map of the whole follow-up system: scan Gmail → build a control sheet →
review → send, looping every 7 days until each reader replies, gets the book,
or opts out.

---

## The pipeline

```
┌─────────────────────────────────────────────────────────────────────┐
│  STAGE 1 — SCAN          response_report.py        (read-only, ~4min) │
│                                                                       │
│   Gmail  ──Sent folder──▶  every message YOU sent (incl. manual       │
│     │                       replies + audiobook links)                │
│     └────Inbox─────────▶  every reply THEY sent                       │
│                                                                       │
│   Detects per person:  replied? · opted-out? · has the book? ·        │
│                        needs help? · full two-way conversation        │
│                                                                       │
│   Writes ▶  responded.csv · no_response.csv ·                         │
│             has_book.txt · needs_help.txt                             │
└─────────────────────────────────────────────────────────────────────┘
                                   │
                                   ▼
┌─────────────────────────────────────────────────────────────────────┐
│  STAGE 2 — BUILD         build_control_sheet.py        (fast)         │
│                                                                       │
│   Reads ▶ the CSVs + has_book/needs_help/suppression.txt +            │
│           followup_sent_log.txt + templates/                          │
│                                                                       │
│   Builds ▶ followups.xlsx                                             │
│     • "Follow-ups" tab — 1 row/person: flags, Conversation Log,       │
│        dropdowns, Follow-ups Sent / History                           │
│     • "Templates" tab  — editable master template wording             │
│                                                                       │
│   ★ AUTO-QUEUE: 7+ days no activity → Follow-up? = Yes (yellow).      │
│     1st nudge = gentle-nudge; after that, Template left BLANK so YOU   │
│     choose. UNLESS has-book / opted-out / suppressed / never.         │
└─────────────────────────────────────────────────────────────────────┘
                                   │
                                   ▼
┌─────────────────────────────────────────────────────────────────────┐
│  STAGE 3 — REVIEW (you, in Excel)                                     │
│                                                                       │
│   Read Conversation Log + flags →  set Follow-up? Yes/No ·            │
│   pick Template OR write Custom Message · edit Templates tab           │
│   →  SAVE & CLOSE                                                      │
└─────────────────────────────────────────────────────────────────────┘
                                   │
                                   ▼
┌─────────────────────────────────────────────────────────────────────┐
│  STAGE 4 — EXECUTE       double-click  Execute Follow-ups.bat          │
│                                                                       │
│   App Password prompt → PREVIEW (dry run) → type SEND to confirm       │
│                                                                       │
│   ⟳ PRE-REFRESH  (re-run Stage 1 + 2 → pull in newly-emailed people)  │
│            │                                                          │
│            ▼                                                          │
│   For each row marked Yes ──▶  [ guardrail gate, below ]              │
│            │                                                          │
│            ▼                                                          │
│   ⟳ POST-REFRESH (re-run Stage 1 + 2 → reflect what just went out)   │
└─────────────────────────────────────────────────────────────────────┘
                                   │
                                   ▼
              ┌──────────── PER-ROW GUARDRAIL GATE ────────────┐
              │                                                 │
   Follow-up? = Yes ?    ──No──▶ skip (ignore row)              │
              │ Yes                                             │
              ▼                                                 │
   Never Follow-up? = yes? ─Yes─▶ SKIPPED (never follow-up)     │
              │ No                                              │
              ▼                                                 │
   In suppression.txt ?  ──Yes─▶ SKIPPED (suppressed)           │
              │ No                                              │
              ▼                                                 │
   In has_book.txt ?     ──Yes─▶ SKIPPED (has the book)         │
              │ No                                              │
              ▼                                                 │
   < 7 days since last                                          │
   follow-up ?           ──Yes─▶ HELD (stays Yes, retries later)│
              │ No                                              │
              ▼                                                 │
   Template already sent to them (Allow Repeat? = no)?          │
     ──Yes─▶ BLOCKED: pick a different template                 │
   Identical custom message already sent?                       │
     ──Yes─▶ BLOCKED: edit the custom message                   │
              │ No                                              │
              ▼                                                 │
   SEND threaded reply (Custom Message if set, else Template)    │
              │                                                 │
              ▼                                                 │
   • log to followup_sent_log.txt (with content fingerprint)    │
   • Follow-ups Sent +1 · append Follow-up History · Status=SENT │
   • Follow-up? → No, Allow Repeat? → no  (both consumed)        │
              └─────────────────────────────────────────────────┘
                                   │
                                   ▼
        ╔══════════════════════════════════════════════════════╗
        ║  THE LOOP:  7 days later → re-run Stage 1 + 2 →       ║
        ║  idle people auto-queue again → repeat until they     ║
        ║  reply · get the book · or opt out                    ║
        ╚══════════════════════════════════════════════════════╝

  ⟳ AUTO-REFRESH: every real send run (Stage 4) automatically re-runs
    Stage 1 (report) + Stage 2 (build) BOTH before sending (to pull in
    anyone newly emailed since last run) AND after (to reflect what just
    went out). So followups.xlsx is current on both ends. Skip with --no-refresh.
```

---

## File legend

| File | Role |
|---|---|
| `response_report.py` | Scans Gmail → CSVs + flag files |
| `build_control_sheet.py` | CSVs → `followups.xlsx` (+ auto-queue) |
| `execute_followups.py` | Sends the marked follow-ups |
| `Execute Follow-ups.bat` | Your double-click "button" |
| `followups.xlsx` | Your control panel (2 tabs) |
| `templates/` | Seed copies of templates (`<name>.txt`) |
| `has_book.txt` / `needs_help.txt` | Auto-detected flags |
| `suppression.txt` | Hard opt-out list |
| `followup_sent_log.txt` | Durable record of every send |

---

## The "Follow-ups" tab columns

`email · responded · Has Book? · Needs Help? · emailed_on · latest_reply_on ·
Never Follow-up? · Follow-up? · Template · Custom Message · Allow Repeat? ·
Follow-ups Sent · Follow-up History · Status · Conversation Log`

- **Has Book?** (peach) — auto-skipped from follow-ups.
- **Needs Help?** (red) — NOT skipped; reach out and assist.
- **Never Follow-up?** (red when `yes`) — manual, permanent hard stop; never
  changed by the program.
- **Follow-up?** (yellow when auto-queued) — `Yes` = send; auto-resets to `No`
  after a send.
- **Template / Custom Message** — Custom wins if both are set. A template the
  person already received is BLOCKED at send time. The auto-queue does NOT pick
  a new template for you — once they've had one, it leaves Template blank so you
  choose the next one (or write a custom). An identical custom message is
  BLOCKED until edited.
- **Allow Repeat?** — set `yes` to override the block and re-send the SAME
  template to someone (explicit, one-time — resets to `no` after the send).
- **Conversation Log** (last column) — full two-way thread, timestamped.

Placeholders usable in any template/custom message:
`{booksprout_link}` · `{got_it_button}` · `{trouble_button}`

---

## Quick command reference

Set credentials once per terminal (App Password, not your normal password):

```powershell
$env:SMTP_USER = "saxena.ankit123@gmail.com"
$env:SMTP_PASS = "your-16-char-app-password"
$py = "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe"
cd "C:\Users\Prisha\OneDrive\Documents\Miner Town\miner_town_followup_kit"
```

| Do this | Command |
|---|---|
| Refresh replies/flags/conversation | `& $py response_report.py` |
| Rebuild the control sheet | `& $py build_control_sheet.py` |
| Preview sends (nothing sent) | `& $py execute_followups.py` |
| Send for real (capped at 40) | `& $py execute_followups.py --send --limit 40` |

Or just double-click **Execute Follow-ups.bat** for the guided
prompt → preview → confirm → send.

---

## Two things to remember

1. **Run Stage 1 (report) before Stage 2 (build)** when you want fresh reply
   and flag data — that's what catches new "got it" / "trouble" signals and
   refreshes the conversation logs.
2. **Stage 4 always previews before sending** and waits for you to type `SEND`,
   so nothing can fire accidentally.
