# Miner Town — Follow-up Nudge Kit

Sends a single, gentle follow-up to everyone who got your original
"MINER TOWN: AWAKENING" email but **never replied** — as a **reply that threads
on the original** (your note on top, your original quoted underneath, so it pops
back to the top of the same Gmail conversation).

People who replied are skipped. People already nudged are skipped. Opt-outs are
honored. Nothing sends until you explicitly pass `--send`.

---

## 1. The files in this kit
| File | What it is | Do you edit it? |
|---|---|---|
| `followup_campaign.py` | The program. Run this. | No |
| `followup_body.txt` | The plain-text follow-up note (sits on top of the quote). | Optional — make it your voice |
| `followup_body.html` | HTML version of the same note. | Optional |
| `suppression.txt` | Addresses to never contact. | Yes — add opt-outs/clients |
| `recipients.txt` | Optional manual list of who you emailed. | Only if NOT using `--from-sent` |
| `README_FOLLOWUP.md` | This guide. | No |
| `followup_sent_log.txt` | Auto-created. The list of who got a follow-up. | No (read it to track) |

You also need **one** source of "who did I email":
- **Easiest:** `--from-sent` reads it straight from your Gmail Sent folder. Nothing to prepare.
- **Or:** a `sent_log.txt` (if you used `send_campaign.py`).
- **Or:** paste the addresses into `recipients.txt`.

---

## 2. One-time setup (5 minutes)
1. Put all the kit files in one folder.
2. **Enable IMAP in Gmail:** Settings (gear) → See all settings → Forwarding and
   POP/IMAP → **Enable IMAP** → Save.
3. **Create a Gmail App Password** (needed instead of your normal password):
   Google Account → Security → 2-Step Verification (must be on) → App passwords →
   create one, copy the 16-character code.
4. (Optional) Edit `followup_body.txt` / `.html` so the note reads in your voice.
   Your BookSprout link is already filled in.

---

## 3. Run it

Open a terminal in the folder and set your login (use the App Password, not your
normal password):

```bash
export SMTP_USER="saxena.ankit123@gmail.com"
export SMTP_PASS="your-16-char-app-password"
```

**Step A — Preview (sends nothing).** Auto-discovers recipients from your Sent folder:

```bash
python3 followup_campaign.py --from-sent
```

It prints who *would* get a nudge, the greeting it'll use, and whether each will
be a true `threaded` reply or a `subject-threaded` "Re:". Read this carefully.

**Step B — Send for real**, in small human-paced batches:

```bash
python3 followup_campaign.py --from-sent --send --min-days 5 --limit 40
```

Re-run the same command another day for the next batch — anyone already nudged is
skipped automatically.

### Using a manual list instead of --from-sent
Paste addresses into `recipients.txt`, then:

```bash
python3 followup_campaign.py --sent-log recipients.txt            # preview
python3 followup_campaign.py --sent-log recipients.txt --send --limit 40
```

### Nicer first-name greetings (optional)
```bash
python3 followup_campaign.py --from-sent --send --csv /path/to/Connections.csv
```

---

## 4. How to tell who got a follow-up
Every send is recorded in **`followup_sent_log.txt`** — a separate file from your
original send record, so the two never mix. It's a CSV you can open in
Excel/Sheets:

```
email,followed_up_at,mode,first_name
reader1@example.com,2026-06-17T09:00:03,threaded,Reader1
reader2@example.com,2026-06-17T09:00:31,subject-threaded,Reader2
```

- **mode** = `threaded` (replied in-thread on your original) or `subject-threaded`
  ("Re:" fallback when the original couldn't be located).
- The script reads this file every run, so nobody is ever nudged twice.

Your three clean lists: who you emailed (Sent folder / `sent_log.txt`), who you
nudged (`followup_sent_log.txt`), who's excluded (`suppression.txt`).

---

## 5. Options
- `--from-sent` — pull the recipient list from your Gmail Sent folder by subject.
- `--min-days N` — only nudge if it's been ≥ N days since the original (default 5).
- `--limit N` — max sends per run (default 50). Keeps you human-paced.
- `--orig-subject "..."` — subject of your original (default "MINER TOWN: AWAKENING").
- `--sent-folder "..."` — Gmail Sent folder (default "[Gmail]/Sent Mail").
- `--sent-log PATH` — a list of prior recipients (e.g. `recipients.txt`).
- `--booksprout-link URL` — override the link (already set to your real one).
- `--csv PATH` — map emails → first names from your LinkedIn export.
- `--send` — actually send (otherwise it's a dry run).
- `--no-imap` — skip reply-check + threading. Not recommended.

---

## 6. Safety / compliance (built in)
- Dry-run by default; nothing sends without `--send`.
- One follow-up per person, ever.
- Honors `suppression.txt`; auto-suppresses anyone who replies declining.
- Skips everyone who replied — no pestering engaged readers.
- Honest-review ask only, with an easy opt-out in every message. Never conditions
  a free copy on a positive review.
- Random 8–22s gap between sends + `--limit` for deliverability.

## 7. Running it on a schedule (optional)
It's a run-once script, not a background service. To repeat it daily, schedule it
(your machine must be awake when it fires):
- **Mac/Linux (cron):**
  `0 9 * * * cd /path/to/folder && SMTP_USER=you@gmail.com SMTP_PASS=app-pass python3 followup_campaign.py --from-sent --send --limit 30 >> followup.log 2>&1`
- **Windows:** Task Scheduler → run `python followup_campaign.py --from-sent --send --limit 30`.
The "one follow-up per person, ever" guard means even a daily run never double-nudges.

## 8. Notes / limits
- Gmail can see who *replied*, not who claimed the ARC or reviewed (that's in
  BookSprout/Goodreads). "No reply after N days" is the nudge signal.
- Personal Gmail caps ~500 emails/day; spread big batches over days with `--limit`.
- If your Sent folder has a non-English name, pass it via `--sent-folder`.
