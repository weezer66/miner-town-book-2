# Miner Town Follow-up Kit — Detailed Program Flow

A deep look at *how the programs actually work* — the data they read/write and
the decisions they make inside each step. (For the short version, see FLOW.md.)

---

## 0. The big picture — who reads/writes what

```
                      ┌───────────────────────────┐
        ┌────────────▶│        Gmail (IMAP)        │
        │ read-only   │  Sent Mail  +  Inbox       │
        │             └───────────────────────────┘
        │                         │ scan
        │                         ▼
 ┌──────────────┐   writes   ┌─────────────────────────────────────────┐
 │ response_    │──────────▶ │ responded.csv  no_response.csv           │
 │ report.py    │            │ has_book.txt   needs_help.txt            │
 └──────────────┘            └─────────────────────────────────────────┘
                                         │ read
                                         ▼
 ┌──────────────┐  also reads ┌─────────────────────────────────────────┐
 │ build_       │◀────────────│ suppression.txt · followup_sent_log.txt  │
 │ control_     │             │ templates/*.txt · existing followups.xlsx│
 │ sheet.py     │  writes     └─────────────────────────────────────────┘
 └──────┬───────┘──────────────────────▶  followups.xlsx  (your control panel)
        │                                          │
        │                                          │ you review & mark
        ▼                                          ▼
 ┌──────────────┐   reads followups.xlsx + the txt logs
 │ execute_     │   sends email via Gmail (SMTP)
 │ followups.py │   appends followup_sent_log.txt, updates followups.xlsx
 └──────────────┘   (and re-runs the two scripts above, before & after)
```

Three programs, run in order. `response_report` and `build_control_sheet`
together = a "refresh." `execute_followups` does the sending (and refreshes
itself on both ends).

---

## 1. response_report.py — read your Gmail, build the picture (READ-ONLY)

Never sends anything. Just reads Gmail and writes data files.

```
START
  │
  ├─ log in to Gmail IMAP  (SMTP_USER + 16-char App Password)
  │
  ├─ scan_sent(): open "[Gmail]/Sent Mail"
  │     SEARCH SUBJECT "MINER TOWN: AWAKENING"  → all thread messages
  │     for each message:
  │        • read headers: To / Cc / Bcc, Date, Subject
  │        • original outreach (exact subject) OR a reply you sent ("Re:")?
  │        • if a reply → fetch only the first ~6 KB of the body, clean it
  │          (skips the big PDF/EPUB attachments on your sends = fast)
  │        • file it under each recipient: (date, snippet, is_original)
  │     ⇒ outbound = { email : [ every message YOU sent them ] }
  │     ⇒ recips   = { email : earliest send date }   (= "emailed_on")
  │
  ├─ for EACH recipient:
  │     reply_status(): open INBOX, SEARCH FROM <recipient>
  │        for each reply (full fetch — their replies are small):
  │           • date + cleaned snippet (their own words, quotes stripped)
  │           • opted_out?  body has "no thanks / unsubscribe / stop / ..."
  │           • has_book?   subject is the "I got it on BookSprout" button,
  │                         OR words like "preordered / downloaded / got it"
  │           • needs_help? subject is the "I'm having trouble" button,
  │                         OR words like "can't download / didn't work"
  │        ⇒ inbound = [ (date, snippet), ... ] + the 3 flags
  │
  │     fmt_conversation(outbound, inbound):
  │        merge your sends (→) and their replies (←), sort by time:
  │           2026-06-10 20:23  →  SENT: original outreach
  │           2026-06-11 01:40  ←  REPLY: tried BookSprout, didn't work
  │           2026-06-11 18:57  →  SENT: here are the PDF/EPUB + audiobook
  │
  └─ WRITE four files:
        responded.csv    repliers + opted_out/has_book/needs_help + conversation
        no_response.csv  people who never replied + their conversation
        has_book.txt     emails to AUTO-SKIP (they have the book)
        needs_help.txt   emails to FLAG (reach out and help)
END
```

---

## 2. build_control_sheet.py — turn the data into your control sheet (fast)

No network. Reads the data files and writes `followups.xlsx`. The clever part
is the **per-row decision** (auto-queue + template choice).

```
START
  ├─ READ: responded.csv · no_response.csv · has_book.txt · needs_help.txt
  │        suppression.txt · followup_sent_log.txt · templates/*.txt
  │        existing followups.xlsx  ← to PRESERVE choices you already made
  │
  ├─ blocked = has_book ∪ opted_out ∪ suppressed   (never auto-queue these)
  │
  ├─ FOR EACH recipient → make one row:
  │   │
  │   │  pull preserved values: Follow-up?, Template, Custom Message,
  │   │     Allow Repeat?, Never Follow-up?, Status
  │   │  received = templates already sent to them (from the log)
  │   │  last_activity = newest timestamp across conversation + log + dates
  │   │
  │   ▼
  │   Never Follow-up? = yes ?
  │      ├─ YES → Follow-up? = No        (hard stop, never queue)
  │      └─ NO  ↓
  │   in `blocked` (has book / opted out / suppressed) ?
  │      ├─ YES → do NOT auto-queue      (keep your prior choice)
  │      └─ NO  ↓
  │   last_activity ≥ 7 days ago ?
  │      ├─ NO  → keep prior choice (usually stays No)
  │      └─ YES → AUTO-QUEUE this person:
  │                 Follow-up? = Yes  (cell highlighted yellow)
  │                 pick the Template:
  │                    • received nothing yet      → gentle-nudge
  │                    • already got a template     → BLANK (you choose)
  │                    • you'd manually picked one
  │                      that isn't sent yet         → keep it
  │   │
  │   └─ also write: Follow-ups Sent (count) · Follow-up History (dated list)
  │                  Has Book? · Needs Help? flags · Conversation Log
  │
  ├─ add dropdowns:  Follow-up? (No/Yes) · Template (names) ·
  │                  Never Follow-up? (no/yes) · Allow Repeat? (no/yes)
  ├─ red highlight rule: Never Follow-up? = yes  → red cell, white text
  ├─ build the "Templates" tab (editable master wording)
  └─ WRITE followups.xlsx   (two tabs: Follow-ups · Templates)
END
```

**You then review the sheet** — read the Conversation Log, flip people to/from
Yes, pick templates or write custom notes, set Never Follow-up? / Allow Repeat?,
save & close.

---

## 3. execute_followups.py — send the marked follow-ups

```
START  (execute_followups.py --send --limit N)
  │
  ├─ [PRE-REFRESH]  (unless --no-refresh)
  │     run response_report.py  +  build_control_sheet.py
  │     → pulls in anyone newly emailed since the last run
  │
  ├─ load followups.xlsx ("Follow-ups" tab)
  ├─ load:  suppression.txt · has_book.txt
  │         last-follow-up time per person   (for cooldown)
  │         sent_history per person: {templates already sent, content hashes}
  │
  ├─ gather every row where Follow-up? = Yes
  │
  ├─ FOR EACH such row  (human-paced; stop after N successful sends):
  │   │
  │   │   ═══════ PER-ROW GUARDRAIL GATE (first match wins) ═══════
  │   │   Never Follow-up? = yes ───────────▶ SKIP  "never follow-up"
  │   │   email in suppression.txt ─────────▶ SKIP  "suppressed"
  │   │   email in has_book.txt ────────────▶ SKIP  "has the book"
  │   │   < 7 days since last follow-up ────▶ HOLD  (stays Yes, retry later)
  │   │
  │   │   choose the body:
  │   │     Custom Message present? → use it (source = "custom")
  │   │     else Template chosen?   → use it (source = template name)
  │   │     else ────────────────────────────▶ ERROR "pick a template/custom"
  │   │
  │   │   repeat guard (fingerprint the rendered text):
  │   │     custom & identical text already sent ─▶ BLOCK "edit the message"
  │   │     template already sent & Allow Repeat?=no ─▶ BLOCK "choose another"
  │   │   ════════════════════════════════════════════════════════
  │   │            │ passes every gate
  │   │            ▼
  │   │   find your original email → build a THREADED reply
  │   │   (your note on top, original quoted below, In-Reply-To/References set)
  │   │            │
  │   │            ▼  SMTP send  ✉
  │   │   on success:
  │   │     • append followup_sent_log.txt:  email,time,source,mode,hash
  │   │     • Follow-ups Sent +1 · Follow-up History += timestamp
  │   │     • Status = "SENT <time> (<source>)"
  │   │     • Follow-up? → No   and   Allow Repeat? → no   (both consumed)
  │   │     • save sheet · wait 8–22 s (look human)
  │   │   on error: print it, mark Status, leave Follow-up? = Yes (retry); go on
  │   │
  │   └─ (HELD / BLOCKED rows stay Yes so you can fix & re-run)
  │
  └─ [POST-REFRESH]  (unless --no-refresh, and only if ≥1 was sent)
        run response_report.py  +  build_control_sheet.py
        → sheet reflects what just went out + any new replies
END
```

---

## 4. The data files (state that lives between runs)

| File | Written by | Holds |
|---|---|---|
| `responded.csv` | report | repliers + flags + full conversation |
| `no_response.csv` | report | silent recipients + their conversation |
| `has_book.txt` | report | auto-skip list (have the book) |
| `needs_help.txt` | report | flag list (stuck — help them) |
| `suppression.txt` | you / opt-outs | hard do-not-contact list |
| `followup_sent_log.txt` | execute | every send ever: email, time, template/custom, content hash |
| `followups.xlsx` | build | your control panel (the only file you edit) |
| `templates/*.txt` | you | seed copies of template wording |

---

## 5. The repeating cycle

```
   you email new people (any time)
            │
            ▼
   ┌───────────────────────────────────────────────┐
   │  REFRESH  =  response_report → build_control_  │  ← also runs automatically
   │  (re-scan Gmail, rebuild the sheet)            │     before & after every send
   └───────────────────────────────────────────────┘
            │
            ▼
   review followups.xlsx (the auto-queue did most of it)
            │
            ▼
   execute_followups --send   (refresh → guardrail gate → send → refresh)
            │
            ▼
   wait ~7 days → idle people auto-queue again → repeat
            │
            ▼
   a person drops out for good when they: reply meaningfully ·
   get the book · opt out · or you mark Never Follow-up? = yes
```
