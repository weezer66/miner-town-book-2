#!/usr/bin/env python3
"""
followup_campaign.py — gentle, threaded follow-ups for the Miner Town ARC outreach.

Companion to send_campaign.py. For everyone you already emailed (from
sent_log.txt), it checks Gmail to see who NEVER replied, then sends each of them
ONE soft follow-up as a REPLY to your original message — "Re: MINER TOWN:
AWAKENING", your new note on top, your original email quoted underneath. That
threads it back to the top of the same conversation in their inbox.

People who replied are skipped. People who already got a follow-up are skipped.
Opt-outs are honored.

WHAT IT CAN / CAN'T SEE
  - CAN (Gmail IMAP): whether a person ever replied to you.
  - CAN'T: whether they claimed the BookSprout ARC or left a review (that lives
    in BookSprout / Goodreads, not your inbox).
  Target = "emailed >= N days ago AND never replied."

SAFETY / COMPLIANCE
  - Dry-run by default. Nothing sends without --send.
  - One follow-up per person, ever (followup_sent_log.txt).
  - Honors suppression.txt; auto-adds anyone whose reply says no/stop/unsubscribe.
  - Honest-review ask only, with an easy opt-out. Never conditions a free copy
    on a positive review.
  - Human-paced: random delay between sends, plus --limit per run.

USAGE
  Preview (sends nothing):
      export SMTP_USER="saxena.ankit123@gmail.com"
      export SMTP_PASS="your-16-char-app-password"
      python3 followup_campaign.py

  Send for real:
      python3 followup_campaign.py --send --limit 40 --min-days 5

  Nicer greetings (emails -> first names):
      python3 followup_campaign.py --send --csv /path/to/Connections.csv

Standard library only.
"""

import argparse
import csv
import email
import imaplib
import os
import random
import re
import smtplib
import ssl
import sys
import time
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import getaddresses, parsedate_to_datetime

# ------------------------------------------------------------------ config
IMAP_HOST = "imap.gmail.com"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465  # SSL

# Without these a half-open socket blocks forever: a scan once hung for 2h43m.
IMAP_TIMEOUT = 120
SMTP_TIMEOUT = 120

ORIGINAL_SUBJECT = "MINER TOWN: AWAKENING"   # used to find your original + as the Re: subject
SENT_FOLDER = "[Gmail]/Sent Mail"            # Gmail's Sent folder (English accounts)
BOOKSPROUT_LINK = "https://booksprout.co/reviewer/review-copy/view/290578/miner-town-awakening"

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
OPT_OUT_WORDS = ("unsubscribe", "no thanks", "no thank you", "stop", "remove me",
                 "opt out", "opt-out", "not interested", "please stop")

# Short follow-up note that sits ON TOP of the quoted original.
DEFAULT_TEXT = """\
Hi {first_name},

Just floating this back to the top of your inbox — no pressure at all. You'd
mentioned you were interested in my novel Miner Town: Awakening, and I know how
easily a note like this slips away.

Your free advance copy is still waiting here:
{booksprout_link}

If you read it, an honest line or two on Goodreads or BookBub means the world for
a debut — but only if you're inclined; the copy is yours either way. Prefer
audio? Just reply "audio". And if you'd rather not hear about it again, reply
"no thanks" and I won't follow up.

Thank you,
Ankit
"""

DEFAULT_HTML = """\
<div style="font-family:Georgia,serif;font-size:16px;line-height:1.55;color:#222;max-width:560px">
  <p>Hi {first_name},</p>
  <p>Just floating this back to the top of your inbox — no pressure at all. You'd
     mentioned you were interested in my novel <i>Miner Town: Awakening</i>, and
     I know how easily a note like this slips away.</p>
  <p>Your free advance copy is still waiting here:<br>
     <a href="{booksprout_link}">{booksprout_link}</a></p>
  <p>If you read it, an honest line or two on Goodreads or BookBub means the
     world for a debut — but only if you're inclined; the copy is yours either
     way. Prefer audio? Just reply <b>"audio"</b>. And if you'd rather not hear
     about it again, reply <b>"no thanks"</b> and I won't follow up.</p>
  <p>Thank you,<br>Ankit</p>
</div>
"""

# Fallback copy of your original, quoted beneath the follow-up if the real sent
# message can't be located in the Sent folder.
EMBEDDED_ORIGINAL_TEXT = """\
MINER TOWN: AWAKENING
Book One of the Miner Town Series by Ankit Saxena

"What will you do to regain the freedom that was once your birthright?"

THE STORY
In a city that runs on Urban Anthracite, the people who mine the fuel keep the
world alive — while the system quietly consumes them. When Gage and Camilla
begin to see it for what it is, they can't unsee it. What follows is a reckoning
with the machine, and the cost of fighting it.

GET YOUR COPY
- Free advance ebook (EPUB or PDF): https://booksprout.co/reviewer/review-copy/view/290578/miner-town-awakening
- Buy or pre-order on Amazon (paperback out July 28): https://a.co/d/0hBRl7oQ
- Audiobook: coming soon to Spotify

READ IT? PLEASE LEAVE AN HONEST REVIEW
Reviews are the single biggest help a book can get. Please post on Goodreads or
BookBub (not Amazon — since I count you as a friend, Amazon doesn't allow your
review there):
- Goodreads: https://www.goodreads.com/book/show/253011399-miner-town
- BookBub: https://www.bookbub.com/books/miner-town-awakening-miner-town-number-1-by-ankit-saxena?source=link_share

Anyone can claim a free copy at: https://ankitsaxenabooks.netlify.app

Thank you for helping the story find its readers.
"""

EMBEDDED_ORIGINAL_HTML = EMBEDDED_ORIGINAL_TEXT.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")


# ------------------------------------------------------------------ helpers
def log(msg):
    print(msg, flush=True)


def imap_arg(s):
    """Quote an IMAP argument (folder name, search string) that may contain
    spaces or special characters. imaplib does not quote these automatically,
    so Gmail rejects e.g. [Gmail]/Sent Mail with 'Could not parse command'."""
    return '"' + str(s).replace('\\', '\\\\').replace('"', '\\"') + '"'


# Lines at/after these markers are quoted history or a signature.
_CUT_MARKERS = re.compile(
    r"^\s*(on .+wrote:|from:\s|sent from my |-{2,}\s*original message|"
    r"_{5,}|get outlook for|>)", re.IGNORECASE)
_INLINE_ATTR = re.compile(r"\bon\s+.{0,120}?\bwrote:.*$", re.IGNORECASE)


def own_words(text):
    """Return just the sender's own words from a reply — quoted original and
    signature stripped. Used so opt-out detection doesn't trip on the quoted
    'reply no thanks' line from our own original email."""
    kept = []
    for line in (text or "").splitlines():
        if _CUT_MARKERS.match(line):
            break
        if line.strip().startswith(">"):
            continue
        kept.append(line.strip())
    snippet = " ".join(w for w in " ".join(kept).split() if w)
    return _INLINE_ATTR.sub("", snippet)


def read_template(path, fallback):
    if path and os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return f.read()
    return fallback


def first_name_from(name, addr):
    if name:
        token = name.strip().split()[0]
        if token and token.lower() not in ("the", "mr", "ms", "dr"):
            return token.capitalize()
    local = addr.split("@")[0]
    local = re.split(r"[._\-+0-9]+", local)[0]
    return local.capitalize() if local else "there"


def load_names_csv(path):
    names = {}
    if not path or not os.path.exists(path):
        return names
    with open(path, encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()
    start = 0
    for i, line in enumerate(lines[:10]):
        if "Email" in line and ("First" in line or "Address" in line):
            start = i
            break
    for row in csv.DictReader(lines[start:]):
        addr = None
        for v in row.values():
            if v and EMAIL_RE.fullmatch(v.strip()):
                addr = v.strip().lower()
                break
        if not addr:
            continue
        fn = (row.get("First Name") or row.get("first_name") or "").strip()
        ln = (row.get("Last Name") or row.get("last_name") or "").strip()
        names[addr] = (fn + " " + ln).strip()
    return names


def parse_dt(token):
    token = token.strip().strip("[]")
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
                "%Y-%m-%d", "%m/%d/%Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(token[:len(fmt) + 4], fmt)
        except ValueError:
            continue
    return None


def load_sent_log(path):
    recips = {}
    if not os.path.exists(path):
        log(f"!! sent_log not found at {path} — nothing to follow up on.")
        return recips
    with open(path, encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            m = EMAIL_RE.search(line)
            if not m:
                continue
            addr = m.group(0).lower()
            dt = None
            for piece in re.split(r"[,\|\t]", line):
                if EMAIL_RE.search(piece):
                    continue
                dt = parse_dt(piece)
                if dt:
                    break
            if addr not in recips or (dt and (recips[addr] is None or dt < recips[addr])):
                recips[addr] = dt
    return recips


def load_set(path):
    out = set()
    if os.path.exists(path):
        with open(path, encoding="utf-8", errors="ignore") as f:
            for line in f:
                m = EMAIL_RE.search(line)
                if m:
                    out.add(m.group(0).lower())
    return out


def append_line(path, text):
    with open(path, "a", encoding="utf-8") as f:
        f.write(text + "\n")


def ensure_followup_header(path):
    """Write a CSV header the first time we log a follow-up."""
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        with open(path, "w", encoding="utf-8") as f:
            f.write("email,followed_up_at,mode,first_name\n")


def extract_part(msg, want):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == want:
                try:
                    return part.get_payload(decode=True).decode(errors="ignore")
                except Exception:
                    continue
        return ""
    try:
        return msg.get_payload(decode=True).decode(errors="ignore")
    except Exception:
        return str(msg.get_payload())


# ------------------------------------------------------------------ Gmail (IMAP)
def imap_connect(user, password):
    M = imaplib.IMAP4_SSL(IMAP_HOST, timeout=IMAP_TIMEOUT)
    M.login(user, password)
    return M


def reply_info(M, addr, since_dt):
    """Returns (replied, opted_out) by looking in INBOX for messages FROM addr."""
    try:
        M.select("INBOX", readonly=True)
        crit = ['FROM', addr]
        if since_dt:
            crit = ['SINCE', since_dt.strftime("%d-%b-%Y"), 'FROM', addr]
        typ, data = M.search(None, *crit)
        if typ != "OK" or not data or not data[0]:
            return False, False
        ids = data[0].split()
        opted = False
        typ, md = M.fetch(ids[-1], "(RFC822)")
        if typ == "OK" and md and md[0]:
            try:
                body = extract_part(email.message_from_bytes(md[0][1]), "text/plain")
                opted = any(w in own_words(body).lower() for w in OPT_OUT_WORDS)
            except Exception:
                pass
        return True, opted
    except Exception as e:
        log(f"   (inbox check failed for {addr}: {e}) — treating as 'no reply'")
        return False, False


def discover_from_sent(M, sent_folder, subject_match, sender):
    """Find everyone you emailed with this subject by scanning Gmail's Sent folder.
    Returns {email: earliest_send_datetime_or_None}."""
    found = {}
    try:
        typ, _ = M.select(imap_arg(sent_folder), readonly=True)
        if typ != "OK":
            log(f"   (could not open Sent folder '{sent_folder}'; try --sent-folder)")
            return found
        typ, data = M.search(None, 'SUBJECT', imap_arg(subject_match))
        if typ != "OK" or not data or not data[0]:
            return found
        for mid in data[0].split():
            typ, md = M.fetch(mid, '(BODY.PEEK[HEADER.FIELDS (TO CC BCC DATE)])')
            if typ != "OK" or not md or not md[0]:
                continue
            hdr = email.message_from_bytes(md[0][1])
            dt = None
            if hdr.get("Date"):
                try:
                    dt = parsedate_to_datetime(hdr.get("Date")).replace(tzinfo=None)
                except Exception:
                    dt = None
            addrs = []
            for h in ("To", "Cc", "Bcc"):
                addrs += [a for _, a in getaddresses(hdr.get_all(h, []))]
            for a in addrs:
                a = a.strip().lower()
                if not a or "@" not in a or a == sender.lower():
                    continue
                if a not in found or (dt and (found[a] is None or dt < found[a])):
                    found[a] = dt
    except Exception as e:
        log(f"   (Sent-folder scan failed: {e})")
    return found


def find_original(M, addr, sent_folder, subject_match):
    """Locate your original sent message to addr; return threading + quote data."""
    try:
        typ, _ = M.select(imap_arg(sent_folder), readonly=True)
        if typ != "OK":
            return None
        typ, data = M.search(None, 'TO', addr, 'SUBJECT', imap_arg(subject_match))
        if typ != "OK" or not data or not data[0]:
            return None
        typ, md = M.fetch(data[0].split()[-1], "(RFC822)")
        if typ != "OK" or not md or not md[0]:
            return None
        msg = email.message_from_bytes(md[0][1])
        return {
            "message_id": msg.get("Message-ID"),
            "references": msg.get("References"),
            "date": msg.get("Date"),
            "text": extract_part(msg, "text/plain") or EMBEDDED_ORIGINAL_TEXT,
            "html": extract_part(msg, "text/html") or EMBEDDED_ORIGINAL_HTML,
        }
    except Exception as e:
        log(f"   (sent-folder lookup failed for {addr}: {e})")
        return None


# ------------------------------------------------------------------ build reply
def build_reply(sender, to_addr, text_tmpl, html_tmpl, fields, orig, subject_base,
                preformatted=False):
    """Build a threaded reply. If preformatted=True, text_tmpl/html_tmpl are
    used as the final body verbatim (no .format) — used for custom messages."""
    em = EmailMessage()
    em["From"] = sender
    em["To"] = to_addr
    em["Subject"] = "Re: " + subject_base

    msg_id = orig.get("message_id") if orig else None
    if msg_id:
        # Collapse any folding whitespace / newlines — email headers must be
        # single-line, and threaded References values often arrive folded.
        msg_id = " ".join(msg_id.split())
        em["In-Reply-To"] = msg_id
        refs = " ".join(((orig.get("references") or "") + " " + msg_id).split())
        em["References"] = refs

    orig_date = orig.get("date") if orig else None
    attribution = (f"On {orig_date}, Ankit Saxena <{sender}> wrote:"
                   if orig_date else f"On an earlier date, Ankit Saxena <{sender}> wrote:")

    q_text = (orig.get("text") if orig else None) or EMBEDDED_ORIGINAL_TEXT
    q_html = (orig.get("html") if orig else None) or EMBEDDED_ORIGINAL_HTML

    body_text = text_tmpl if preformatted else text_tmpl.format(**fields)
    quoted_text = "\n".join("> " + ln for ln in q_text.splitlines())
    em.set_content(f"{body_text}\n\n{attribution}\n{quoted_text}\n")

    body_html = html_tmpl if preformatted else html_tmpl.format(**fields)
    full_html = (
        f"{body_html}"
        f'<br><div style="color:#666;font-size:13px">{attribution}</div>'
        f'<blockquote style="margin:6px 0 0 .8ex;border-left:2px solid #ccc;'
        f'padding-left:12px;color:#555">{q_html}</blockquote>'
    )
    em.add_alternative(full_html, subtype="html")
    return em


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description="One-time gentle, threaded follow-ups for non-repliers.")
    ap.add_argument("--sent-log", default="sent_log.txt")
    ap.add_argument("--suppression", default="suppression.txt")
    ap.add_argument("--followup-log", default="followup_sent_log.txt")
    ap.add_argument("--csv", default=None, help="Optional Connections.csv to map emails -> first names.")
    ap.add_argument("--text", default="followup_body.txt")
    ap.add_argument("--html", default="followup_body.html")
    ap.add_argument("--booksprout-link", default=os.environ.get("BOOKSPROUT_LINK", BOOKSPROUT_LINK))
    ap.add_argument("--sent-folder", default=SENT_FOLDER, help="Gmail Sent folder name.")
    ap.add_argument("--orig-subject", default=ORIGINAL_SUBJECT, help="Subject of your original email (to find + thread on).")
    ap.add_argument("--from-sent", action="store_true",
                    help="Discover recipients by scanning your Gmail Sent folder for the original subject (no sent_log needed).")
    ap.add_argument("--min-days", type=int, default=5)
    ap.add_argument("--limit", type=int, default=50)
    ap.add_argument("--send", action="store_true", help="Actually send. Without this it's a dry run.")
    ap.add_argument("--no-imap", action="store_true", help="Skip reply check AND threading (not recommended).")
    args = ap.parse_args()

    user = os.environ.get("SMTP_USER")
    pw = os.environ.get("SMTP_PASS")
    if not user or not pw:
        sys.exit("Set SMTP_USER and SMTP_PASS (a Gmail App Password) as environment variables first.")

    text_tmpl = read_template(args.text, DEFAULT_TEXT)
    html_tmpl = read_template(args.html, DEFAULT_HTML)

    recips = load_sent_log(args.sent_log)
    suppressed = load_set(args.suppression)
    already = load_set(args.followup_log)
    names = load_names_csv(args.csv)

    log(f"Loaded {len(recips)} prior recipients, {len(suppressed)} suppressed, "
        f"{len(already)} already followed up.")
    log(f"Mode: {'SEND' if args.send else 'DRY RUN (no emails will be sent)'} · "
        f"min-days={args.min_days} · limit={args.limit}\n")

    M = None
    if not args.no_imap:
        try:
            M = imap_connect(user, pw)
        except Exception as e:
            sys.exit(f"Could not connect to Gmail IMAP: {e}\n"
                     f"(Enable IMAP in Gmail settings and use an App Password.)")

    smtp = None
    if args.send:
        smtp = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT,
                                context=ssl.create_default_context())
        smtp.login(user, pw)

    if args.from_sent:
        if M is None:
            sys.exit("--from-sent needs Gmail access; remove --no-imap.")
        disc = discover_from_sent(M, args.sent_folder, args.orig_subject, user)
        log(f"Discovered {len(disc)} recipient(s) from Sent folder matching subject '{args.orig_subject}'.")
        for a, d in disc.items():
            if a not in recips or (d and (recips[a] is None or d < recips[a])):
                recips[a] = d
        log(f"Total recipients to consider: {len(recips)}\n")

    now = datetime.now()
    sent = skipped_replied = skipped_recent = skipped_supp = skipped_done = 0
    candidates = []
    for addr, sent_dt in recips.items():
        if addr in suppressed:
            skipped_supp += 1; continue
        if addr in already:
            skipped_done += 1; continue
        if sent_dt and (now - sent_dt) < timedelta(days=args.min_days):
            skipped_recent += 1; continue
        candidates.append((addr, sent_dt))
    candidates.sort(key=lambda x: (x[1] or datetime.min))  # oldest first

    for addr, sent_dt in candidates:
        if sent >= args.limit:
            log(f"\nReached --limit ({args.limit}). Re-run to continue.")
            break

        orig = None
        if M is not None:
            replied, opted = reply_info(M, addr, sent_dt)
            if opted:
                append_line(args.suppression, addr)
                log(f"[opt-out]  {addr} — replied declining; suppressed, skipping.")
                skipped_replied += 1; continue
            if replied:
                log(f"[replied]  {addr} — already responded, skipping.")
                skipped_replied += 1; continue
            orig = find_original(M, addr, args.sent_folder, args.orig_subject)

        fields = {
            "first_name": first_name_from(names.get(addr, ""), addr),
            "booksprout_link": args.booksprout_link,
        }
        mode = "threaded" if (orig and orig.get("message_id")) else "subject-threaded"

        if not args.send:
            log(f"[would send] {addr}  ({fields['first_name']}, {mode})")
            sent += 1; continue

        try:
            smtp.send_message(build_reply(user, addr, text_tmpl, html_tmpl, fields, orig, args.orig_subject))
            ensure_followup_header(args.followup_log)
            append_line(args.followup_log,
                        f"{addr},{now.isoformat(timespec='seconds')},{mode},{fields['first_name']}")
            log(f"[sent]     {addr}  ({mode})")
            sent += 1
            time.sleep(random.uniform(8, 22))
        except Exception as e:
            log(f"[ERROR]    {addr}: {e}")

    if smtp:
        smtp.quit()
    if M:
        try:
            M.logout()
        except Exception:
            pass

    log("\n--- summary ---")
    log(f"{'Sent' if args.send else 'Would send'}: {sent}")
    log(f"Skipped — replied/opted out: {skipped_replied}")
    log(f"Skipped — too recent (< {args.min_days}d): {skipped_recent}")
    log(f"Skipped — already followed up: {skipped_done}")
    log(f"Skipped — suppressed: {skipped_supp}")
    if args.send:
        log(f"\nFollow-up recipients logged to: {args.followup_log}  (columns: email, followed_up_at, mode, first_name)")
    if not args.send:
        log("\nThis was a DRY RUN. Re-run with --send to actually send.")


if __name__ == "__main__":
    main()
