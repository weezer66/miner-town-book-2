#!/usr/bin/env python3
"""
execute_followups.py — send the follow-ups you marked in followups.xlsx.

Reads followups.xlsx. For every row where Follow-up? = Yes, it sends a threaded
reply to your original MINER TOWN email, then writes the result into the Status
column. The body is your Custom Message if you typed one, otherwise the chosen
Template (custom wins when both are present). Template text is read from the
'Templates' tab of the workbook, so edits you make there take effect.

  DRY RUN (default, sends nothing, just previews):
      python execute_followups.py

  SEND FOR REAL:
      python execute_followups.py --send

After each successful send the row's 'Follow-ups Sent' count is incremented, the
timestamp is appended to 'Follow-up History', and Follow-up? is flipped back to
No so the row is NOT re-sent next run. To follow up with someone again later,
set Follow-up? = Yes again — repeat follow-ups are allowed and each is tallied.

Safety:
  - Dry-run unless --send.
  - Skips anyone in suppression.txt or has_book.txt.
  - Cooldown: HOLDS a send if it's been < --min-gap-days (default 7) since the
    last follow-up to that person. Held rows stay marked Yes and go out on a
    later run once enough time has passed — stops over-eager repeat nudges.
  - Follow-up? auto-resets to No after a send (no accidental repeat on re-run).
  - No repeats: BLOCKS re-sending a template this person already received
    (override per-row with 'Allow Repeat?' = yes, which is consumed after use),
    and BLOCKS an identical custom message (edit the text first — no override).
  - Stop chasing: after the final nudge (the 'still-interested' template) is
    sent, the person is marked Never Follow-up? = yes so they're never queued
    again. Sequence: gentle-nudge -> still-interested -> done.
  - Human-paced delay between sends; --limit caps per run.

A --send run REFRESHES the sheet BEFORE and AFTER sending: it re-runs
response_report.py + build_control_sheet.py first (so anyone newly emailed since
the last run is pulled in and queued), then sends, then refreshes again (so the
sheet reflects what just went out plus any new replies). Pass --no-refresh to
skip both.

Reuses the Gmail/threading logic from followup_campaign.py.
Needs: openpyxl
"""
import argparse
import hashlib
import os
import random
import smtplib
import ssl
import sys
import time
from datetime import datetime, timedelta

from openpyxl import load_workbook

from followup_campaign import (
    imap_connect, find_original, build_reply,
    load_set, append_line,
    SENT_FOLDER, ORIGINAL_SUBJECT, BOOKSPROUT_LINK, SMTP_HOST, SMTP_PORT,
    SMTP_TIMEOUT,
)
from send_guardrail import enforce_send_run_cooldown, mark_send_run_finished

HERE = os.path.dirname(os.path.abspath(__file__))
SHEET = os.path.join(HERE, "followups.xlsx")
TEMPLATES_DIR = os.path.join(HERE, "templates")
SUPPRESSION = os.path.join(HERE, "suppression.txt")
FOLLOWUP_LOG = os.path.join(HERE, "followup_sent_log.txt")
HAS_BOOK_FILE = os.path.join(HERE, "has_book.txt")
SEND_COOLDOWN_LOCK = os.path.join(HERE, "send_cooldown_lock.json")

# After sending this template (the last one in the nudge sequence), the person
# is marked Never Follow-up? = yes — we stop chasing them for good.
STOP_AFTER_TEMPLATE = "still-interested"


def ensure_log_header():
    """Write the durable per-send log header the first time we log anything.
    One line per follow-up ever sent (repeat sends append more lines)."""
    if not os.path.exists(FOLLOWUP_LOG) or os.path.getsize(FOLLOWUP_LOG) == 0:
        with open(FOLLOWUP_LOG, "w", encoding="utf-8") as f:
            f.write("email,sent_at,source,mode,body_hash\n")


def body_fingerprint(text):
    """Short stable hash of a message body — used to detect identical content
    being re-sent (esp. custom messages that weren't edited)."""
    norm = " ".join((text or "").split()).lower()
    return hashlib.sha1(norm.encode("utf-8")).hexdigest()[:12]


def load_sent_history(path):
    """email -> {'templates': set(names already sent), 'hashes': set(body hashes)}."""
    hist = {}
    if not os.path.exists(path):
        return hist
    with open(path, encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line or line.lower().startswith("email,"):
                continue
            parts = [p.strip() for p in line.split(",")]
            if len(parts) < 3:
                continue
            email = parts[0].lower()
            source = parts[2]
            bhash = parts[4] if len(parts) > 4 else ""
            d = hist.setdefault(email, {"templates": set(), "hashes": set()})
            if source and source != "custom":
                d["templates"].add(source)
            if bhash:
                d["hashes"].add(bhash)
    return hist


def refresh_sheet(tag):
    """Re-scan Gmail (response_report) and rebuild the sheet (build_control_sheet)
    so followups.xlsx is current. Returns True on success. Non-fatal on failure."""
    print("\n" + "=" * 60)
    print(f"{tag}: re-scanning Gmail and rebuilding the sheet")
    print("(re-runs the report — a few minutes — then rebuilds)")
    print("=" * 60, flush=True)
    try:
        os.chdir(HERE)               # report writes CSVs relative to CWD
        import response_report
        import build_control_sheet
        response_report.main()
        build_control_sheet.main()
        print(f"\n{tag} complete — followups.xlsx is up to date.")
        return True
    except SystemExit:
        raise
    except Exception as e:
        print(f"\n!! {tag} failed: {e}")
        print("   Refresh manually: response_report.py then build_control_sheet.py.")
        return False


def load_last_followup(path):
    """email -> datetime of the MOST RECENT follow-up already sent, from the log.
    Used to enforce a minimum gap between follow-ups to the same person."""
    last = {}
    if not os.path.exists(path):
        return last
    with open(path, encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line or line.lower().startswith("email,"):
                continue
            parts = [p.strip() for p in line.split(",")]
            if len(parts) < 2:
                continue
            email = parts[0].lower()
            dt = None
            for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d"):
                try:
                    dt = datetime.strptime(parts[1], fmt); break
                except ValueError:
                    continue
            if dt and (email not in last or dt > last[email]):
                last[email] = dt
    return last


TEMPLATES_SHEET = "Templates"

# One-click reply buttons. Drop {got_it_button} / {trouble_button} into any
# template or custom message. They become styled buttons in HTML and labeled
# links in plain text; clicking opens the reader's mail app with a reply to you.
GOT_IT_SUBJECT = "I got the book on BookSprout"
GOT_IT_BODY = ("Hi Ankit,\n\nJust letting you know I downloaded Miner Town: "
               "Awakening from BookSprout. Thank you!")
TROUBLE_SUBJECT = "Trouble with BookSprout"
TROUBLE_BODY = ("Hi Ankit,\n\nI'm having trouble getting Miner Town: Awakening "
                "from BookSprout. Could you help, or send the file directly?\n\nThanks!")


def _mailto(to, subject, body):
    from urllib.parse import quote
    return f"mailto:{to}?subject={quote(subject)}&body={quote(body)}"


def _button_html(url, label, color):
    href = url.replace("&", "&amp;")
    return (f'<a href="{href}" style="display:inline-block;background:{color};'
            f'color:#ffffff;text-decoration:none;padding:10px 18px;border-radius:6px;'
            f'font-family:Arial,sans-serif;font-size:14px;font-weight:bold;'
            f'margin:6px 8px 6px 0">{label}</a>')


def _esc(s):
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return s.replace(BOOKSPROUT_LINK, f'<a href="{BOOKSPROUT_LINK}">{BOOKSPROUT_LINK}</a>')


def text_to_html(text):
    """Turn a plain-text message (template or custom) into simple, safe HTML:
    blank lines start new paragraphs. Within a paragraph, a line break that
    follows a long line is treated as a cosmetic word-wrap (joined with a
    space) so sentences flow; a break after a short line (a link line, a
    signature) is kept as a real <br>. Button tokens pass through untouched
    and are swapped for real button HTML afterwards by render_bodies()."""
    WRAP = 50   # lines at least this long are assumed to be soft-wrapped
    out = []
    for para in (text or "").split("\n\n"):
        para = para.strip("\n")
        if not para.strip():
            continue
        lines = para.split("\n")
        pieces = [_esc(lines[0])]
        for i in range(1, len(lines)):
            sep = " " if len(lines[i - 1].rstrip()) >= WRAP else "<br>"
            pieces.append(sep + _esc(lines[i]))
        out.append(f"<p>{''.join(pieces)}</p>")
    return ('<div style="font-family:Georgia,serif;font-size:16px;line-height:1.55;'
            'color:#222">' + "".join(out) + "</div>")


def render_bodies(raw, user):
    """Expand all placeholders in a template/custom body and return
    (text_body, html_body). user is the From address the buttons reply to."""
    got_url = _mailto(user, GOT_IT_SUBJECT, GOT_IT_BODY)
    trouble_url = _mailto(user, TROUBLE_SUBJECT, TROUBLE_BODY)

    text = raw.replace("{booksprout_link}", BOOKSPROUT_LINK)
    text = text.replace("{got_it_button}",
                        f"✔ I got it on BookSprout — click to tell me: {got_url}")
    text = text.replace("{trouble_button}",
                        f"⚠ Having trouble with BookSprout? Click for help: {trouble_url}")

    html = text_to_html(raw.replace("{booksprout_link}", BOOKSPROUT_LINK))
    html = html.replace("{got_it_button}",
                        _button_html(got_url, "✅ I got it on BookSprout", "#2e7d32"))
    html = html.replace("{trouble_button}",
                        _button_html(trouble_url, "⚠️ I'm having trouble", "#b26a00"))
    return text, html


def load_template_bodies(wb):
    """name -> text. Prefers the workbook's 'Templates' tab; falls back to
    templates/*.txt on disk if that tab is absent (older sheet)."""
    bodies = {}
    if TEMPLATES_SHEET in wb.sheetnames:
        tws = wb[TEMPLATES_SHEET]
        for row in tws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            name = str(row[0]).strip()
            if name == "(read me)":
                continue
            bodies[name] = str(row[1]) if len(row) > 1 and row[1] is not None else ""
        return bodies
    import glob
    for p in sorted(glob.glob(os.path.join(TEMPLATES_DIR, "*.txt"))):
        name = os.path.splitext(os.path.basename(p))[0]
        with open(p, encoding="utf-8") as f:
            bodies[name] = f.read()
    return bodies


def header_map(ws):
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v:
            cols[str(v).strip()] = c
    required = ["email", "Follow-up?", "Template", "Status"]
    missing = [r for r in required if r not in cols]
    if missing:
        sys.exit(f"followups.xlsx is missing column(s): {missing}. Re-run build_control_sheet.py.")
    return cols


def main():
    ap = argparse.ArgumentParser(description="Send the follow-ups marked in followups.xlsx.")
    ap.add_argument("--send", action="store_true", help="Actually send. Without this it's a dry run.")
    ap.add_argument("--limit", type=int, default=50, help="Max sends per run.")
    ap.add_argument("--min-gap-days", type=int, default=7,
                    help="Minimum days since the last follow-up before sending another.")
    ap.add_argument("--no-refresh", action="store_true",
                    help="Skip the automatic report + sheet rebuild after sending.")
    ap.add_argument("--run-cooldown-minutes", type=int, default=10,
                    help="Minimum minutes between any two send runs (default 10).")
    ap.add_argument("--offline-preview", action="store_true",
                    help="Dry-run preview without any Gmail IMAP/SMTP connections.")
    ap.add_argument("--sent-folder", default=SENT_FOLDER)
    ap.add_argument("--orig-subject", default=ORIGINAL_SUBJECT)
    args = ap.parse_args()

    if args.offline_preview and args.send:
        sys.exit("--offline-preview can only be used without --send.")

    user = os.environ.get("SMTP_USER")
    pw = os.environ.get("SMTP_PASS")
    if not user or not pw:
        sys.exit("Set SMTP_USER and SMTP_PASS (Gmail App Password) first.")

    if args.send:
        enforce_send_run_cooldown(
            SEND_COOLDOWN_LOCK,
            args.run_cooldown_minutes,
            "execute_followups.py",
        )

    # PRE-SEND refresh: pull in anyone newly emailed since the last run (and
    # fresh replies/flags) so they're queued before we send. Only on --send.
    if args.send and not args.no_refresh:
        refresh_sheet("PRE-SEND REFRESH")

    if not os.path.exists(SHEET):
        sys.exit("followups.xlsx not found. Run build_control_sheet.py first.")

    try:
        wb = load_workbook(SHEET)
    except PermissionError:
        sys.exit("Could not open followups.xlsx — please CLOSE it in Excel and try again.")
    ws = wb["Follow-ups"] if "Follow-ups" in wb.sheetnames else wb.active
    cols = header_map(ws)
    template_bodies = load_template_bodies(wb)

    suppressed = load_set(SUPPRESSION)
    has_book = load_set(HAS_BOOK_FILE)
    last_followup = load_last_followup(FOLLOWUP_LOG)
    sent_history = load_sent_history(FOLLOWUP_LOG)
    has_tally = "Follow-ups Sent" in cols and "Follow-up History" in cols

    # Gather the rows the user marked Yes.
    has_custom = "Custom Message" in cols
    has_never = "Never Follow-up?" in cols
    has_repeat = "Allow Repeat?" in cols
    chosen = []
    for r in range(2, ws.max_row + 1):
        email = ws.cell(row=r, column=cols["email"]).value
        if not email:
            continue
        email = str(email).strip().lower()
        followup = (ws.cell(row=r, column=cols["Follow-up?"]).value or "").strip()
        if followup != "Yes":
            continue
        never = ""
        if has_never:
            never = (ws.cell(row=r, column=cols["Never Follow-up?"]).value or "").strip().lower()
        allow_repeat = ""
        if has_repeat:
            allow_repeat = (ws.cell(row=r, column=cols["Allow Repeat?"]).value or "").strip().lower()
        template = (ws.cell(row=r, column=cols["Template"]).value or "").strip()
        custom = ""
        if has_custom:
            custom = (ws.cell(row=r, column=cols["Custom Message"]).value or "").strip()
        chosen.append((r, email, template, custom, never, allow_repeat))

    print(f"Marked for follow-up: {len(chosen)}")
    print(f"Mode: {'SEND' if args.send else 'DRY RUN (nothing will be sent)'} · limit={args.limit}\n")
    if not chosen:
        print("Nothing marked 'Yes'. Open followups.xlsx, set Follow-up? = Yes, save, and re-run.")
        return

    # Connect once unless this is an offline dry-run preview.
    M = None
    if not args.offline_preview:
        try:
            M = imap_connect(user, pw)
        except Exception as e:
            sys.exit(f"Could not connect to Gmail IMAP: {e}")

    smtp = None
    if args.send:
        try:
            smtp = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT,
                                    context=ssl.create_default_context())
            smtp.login(user, pw)
        except Exception as e:
            try:
                M.logout()
            except Exception:
                pass
            sys.exit(f"Could not connect to Gmail SMTP: {e}")

    sent = 0
    counts = {"sent": 0, "skipped": 0, "held": 0, "blocked": 0, "error": 0,
              "would_send": 0}

    def set_status(row, text):
        ws.cell(row=row, column=cols["Status"], value=text)

    def bump_tally(row, stamp, source):
        """Increment the per-person send count and append a timestamped line."""
        if not has_tally:
            return
        cnt_cell = ws.cell(row=row, column=cols["Follow-ups Sent"])
        try:
            cnt = int(cnt_cell.value or 0)
        except (TypeError, ValueError):
            cnt = 0
        cnt_cell.value = cnt + 1
        hist_cell = ws.cell(row=row, column=cols["Follow-up History"])
        line = f"{stamp} — {source}"
        hist_cell.value = (str(hist_cell.value) + "\n" + line) if hist_cell.value else line

    def consume(row):
        """Reset Follow-up? to No (and Allow Repeat? to no) after a send so the
        row is not re-sent on the next run and the override is one-time."""
        ws.cell(row=row, column=cols["Follow-up?"], value="No")
        if has_repeat:
            ws.cell(row=row, column=cols["Allow Repeat?"], value="no")

    fields = {"first_name": "", "booksprout_link": BOOKSPROUT_LINK}

    for row, email, template, custom, never, allow_repeat in chosen:
        if sent >= args.limit:
            print(f"\nReached --limit ({args.limit}). Re-run to continue.")
            break

        if never == "yes":
            print(f"[skip] {email} — Never Follow-up? = yes.")
            set_status(row, "SKIPPED (never follow-up)"); counts["skipped"] += 1; continue
        if email in suppressed:
            print(f"[skip] {email} — suppressed.")
            set_status(row, "SKIPPED (suppressed)"); counts["skipped"] += 1; continue
        if email in has_book:
            print(f"[skip] {email} — already has the book.")
            set_status(row, "SKIPPED (has the book)")
            consume(row); counts["skipped"] += 1; continue

        # Cooldown: hold if it's been < min-gap-days since the last follow-up.
        # Stays queued (Follow-up? left = Yes) so it sends once enough time passes.
        last = last_followup.get(email)
        if last is not None:
            age = (datetime.now() - last).days
            if age < args.min_gap_days:
                print(f"[hold] {email} — only {age}d since last follow-up "
                      f"(need {args.min_gap_days}).")
                set_status(row, f"HELD ({age}d since last; need {args.min_gap_days}d) "
                                f"as of {datetime.now():%Y-%m-%d}")
                counts["held"] += 1; continue

        # A custom message wins over a template; otherwise a template is required.
        if custom:
            raw = custom
            source = "custom"
        elif template:
            if template not in template_bodies:
                print(f"[error] {email} — template '{template}' not in the Templates tab.")
                set_status(row, f"ERROR: template '{template}' missing"); counts["error"] += 1; continue
            raw = template_bodies[template]
            source = template
        else:
            print(f"[error] {email} — Yes but no Template chosen and no Custom Message.")
            set_status(row, "ERROR: pick a Template or write a Custom Message")
            counts["error"] += 1; continue

        text_body, html_body = render_bodies(raw, user)

        # Don't repeat content: identical custom message (must be edited), or a
        # template this person already received (unless Allow Repeat? = yes).
        bhash = body_fingerprint(text_body)
        hist = sent_history.get(email, {"templates": set(), "hashes": set()})
        if source == "custom" and bhash in hist["hashes"]:
            print(f"[blocked] {email} — custom message identical to a previous send.")
            set_status(row, "BLOCKED: edit the custom message (same content already sent)")
            counts["blocked"] += 1; continue
        if source != "custom" and source in hist["templates"] and allow_repeat != "yes":
            print(f"[blocked] {email} — '{source}' already sent; choose another template.")
            set_status(row, f"BLOCKED: '{source}' already sent — pick a different "
                            f"template, or set Allow Repeat? = yes")
            counts["blocked"] += 1; continue

        if not args.send:
            if args.offline_preview:
                mode = "offline-preview"
            else:
                orig = find_original(M, email, args.sent_folder, args.orig_subject)
                mode = "threaded" if (orig and orig.get("message_id")) else "subject-threaded"
            print(f"[would send] {email}  (source={source}, {mode})")
            counts["would_send"] += 1; sent += 1; continue

        try:
            orig = find_original(M, email, args.sent_folder, args.orig_subject)
            mode = "threaded" if (orig and orig.get("message_id")) else "subject-threaded"
            smtp.send_message(build_reply(user, email, text_body, html_body,
                                          fields, orig, args.orig_subject, preformatted=True))
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
            ensure_log_header()
            append_line(FOLLOWUP_LOG, f"{email},{stamp},{source},{mode},{bhash}")
            bump_tally(row, stamp, source)
            set_status(row, f"SENT {stamp} ({source})")
            consume(row)                 # flip Follow-up? back to No
            # After the final nudge (the 'still-interested' template), stop
            # chasing this person for good — set Never Follow-up? = yes.
            if source == STOP_AFTER_TEMPLATE and has_never:
                ws.cell(row=row, column=cols["Never Follow-up?"], value="yes")
                print(f"       {email} -> Never Follow-up? = yes "
                      f"(final '{STOP_AFTER_TEMPLATE}' nudge sent)")
            print(f"[sent] {email}  (source={source}, {mode})")
            counts["sent"] += 1; sent += 1
            wb.save(SHEET)               # persist status/tally as we go
            time.sleep(random.uniform(8, 22))
        except Exception as e:
            print(f"[ERROR] {email}: {e}")
            set_status(row, f"ERROR: {e}"); counts["error"] += 1

    if smtp:
        try:
            smtp.quit()
        except Exception:
            pass          # connection may already be dead — don't crash cleanup
    try:
        M.logout()
    except Exception:
        pass

    if args.send:
        try:
            wb.save(SHEET)
        except PermissionError:
            print("\n!! Could not save followups.xlsx (is it open in Excel?). "
                  "Statuses are still recorded in followup_sent_log.txt.")

    print("\n--- summary ---")
    if args.send:
        print(f"Sent: {counts['sent']}")
    else:
        print(f"Would send: {counts['would_send']}")
    print(f"Skipped: {counts['skipped']}")
    print(f"Held (cooldown < {args.min_gap_days}d): {counts['held']}")
    print(f"Blocked (repeat content): {counts['blocked']}")
    print(f"Errors:  {counts['error']}")
    if args.send:
        mark_send_run_finished(SEND_COOLDOWN_LOCK, counts["sent"])
    if counts["blocked"]:
        print("  (blocked rows stay marked Yes — pick a different template, set "
              "Allow Repeat? = yes, or edit the custom message, then re-run.)")
    if counts["held"]:
        print(f"  (held rows stay marked Yes — they'll send once {args.min_gap_days} "
              f"days have passed since their last follow-up.)")
    if not args.send:
        print("\nThis was a DRY RUN. Re-run with --send to actually send.")

    # POST-SEND refresh: re-scan Gmail and rebuild so the sheet reflects what
    # just went out (and any new replies). Skip if nothing was sent.
    if args.send and counts["sent"] > 0 and not args.no_refresh:
        refresh_sheet("POST-SEND AUTO-REFRESH")
    elif args.send and counts["sent"] == 0 and not args.no_refresh:
        print("\n(Nothing sent — skipping post-send refresh.)")


if __name__ == "__main__":
    main()
